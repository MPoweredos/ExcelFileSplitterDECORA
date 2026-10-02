using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed class ExcelSession : IDisposable
{
    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);

    private const uint WmNull = 0x0000;
    private const uint SmtoAbortIfHung = 0x0002;
    private const uint ResponseTimeoutMilliseconds = 5_000;

    private readonly int _excelProcessId;
    private readonly IntPtr _excelWindow;
    private readonly Process? _excelProcess;
    private bool _disposed;

    public Excel.Application App { get; }

    public int ProcessId => _excelProcessId;

    public bool MacrosBlocked { get; }

    public TimeSpan Startup { get; }

    public TimeSpan Shutdown { get; private set; }

    public TimeSpan Quitting { get; private set; }

    public int? ExitCode { get; private set; }

    public bool Crashed => ExitCode is int code && IsCrash(code);

    public static bool IsCrash(int exitCode) => exitCode is not 0 and not -1;

    public string ExitCodeText() => ExitCode switch
    {
        null => "?",
        < 0 and var code => $"0x{code:X8}",
        var code => code.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    public ExcelSession()
    {
        var startup = Stopwatch.StartNew();
        App = StartExcel();

        _excelWindow = Com.Or(() => new IntPtr(App.Hwnd), IntPtr.Zero);
        _excelProcessId = ProcessOf(_excelWindow);
        ExcelProcessRegistry.Register(_excelProcessId);
        _excelProcess = Watch(_excelProcessId);

        try
        {
            App.Visible = false;
            App.DisplayAlerts = false;
            App.ScreenUpdating = false;
            App.EnableEvents = false;
            App.AskToUpdateLinks = false;
            App.AlertBeforeOverwriting = false;

            MacrosBlocked = Com.Try(() => ((dynamic)App).AutomationSecurity = 3);
        }
        catch
        {
            _disposed = true;
            if (_excelProcessId > 0 && ExcelProcessRegistry.TryKill(_excelProcessId))
                ExcelProcessRegistry.Unregister(_excelProcessId);
            _excelProcess?.Dispose();
            ReleaseComObject(App);
            throw;
        }

        Startup = startup.Elapsed;
    }

    private static int ProcessOf(IntPtr window)
    {
        if (window == IntPtr.Zero) return 0;
        GetWindowThreadProcessId(window, out int processId);
        return processId;
    }

    private static Process? Watch(int processId)
    {
        if (processId <= 0) return null;

        try
        {
            var process = Process.GetProcessById(processId);
            _ = process.SafeHandle;
            return process;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException
                                              or System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return null;
        }
    }

    private static int? ExitCodeOf(Process? process)
    {
        if (process is null) return null;

        try
        {
            return process.HasExited ? process.ExitCode : null;
        }
        catch (Exception exception) when (exception is InvalidOperationException
                                              or System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return null;
        }
    }

    public string DescribeShutdown()
    {
        var parts = new List<string> { $"sam proces {Shutdown.Text()}" };
        if (Quitting >= TimeSpan.FromSeconds(1)) parts.Add($"z czego Quit {Quitting.Text()}");

        if (Crashed) parts.Add($"awaria, kod {ExitCodeText()}");

        return string.Join(", ", parts);
    }

    private bool IsResponding()
    {
        if (_excelWindow == IntPtr.Zero) return true;

        IntPtr sent = SendMessageTimeout(
            _excelWindow, WmNull, IntPtr.Zero, IntPtr.Zero, SmtoAbortIfHung, ResponseTimeoutMilliseconds, out _);
        return sent != IntPtr.Zero;
    }

    private static Excel.Application StartExcel()
    {
        try
        {
            return new Excel.Application();
        }
        catch (Exception exception) when (Com.SaysNo(exception))
        {
            throw new InvalidOperationException(
                "Nie udało się uruchomić Excela. Ten program steruje prawdziwym Excelem, więc musi on być " +
                "zainstalowany na tym komputerze - wersja na komputer (Microsoft 365 albo Office 2016 i nowszy). " +
                "Excel w przeglądarce nie wystarczy. " +
                $"Odpowiedź systemu: {exception.Message}", exception);
        }
    }

    public void ReportStartupProblems(IProgressReporter reporter)
    {
        if (!MacrosBlocked)
            reporter.Warning("nie udało się zablokować makr - skoroszyt .xlsm może uruchomić własne makra " +
                             "przy każdym otwarciu (Workbook_Open, Auto_Open)");
    }

    public Excel.Workbook OpenWorkbook(string path, bool readOnly = false)
    {
        return App.Workbooks.Open(
            Filename: Path.GetFullPath(path),
            UpdateLinks: 0,
            ReadOnly: readOnly,
            IgnoreReadOnlyRecommended: true,
            Notify: false,
            AddToMru: false);
    }

    public T WithWorkbook<T>(string path, bool readOnly, Func<Excel.Workbook, T> work)
    {
        Excel.Workbook? workbook = null;
        try
        {
            workbook = OpenWorkbook(path, readOnly);
            return work(workbook);
        }
        finally
        {
            CloseWorkbook(workbook, save: false);
        }
    }

    public void RefreshAllAndWait(Excel.Workbook workbook, bool refreshPivotCaches, IProgressReporter reporter)
    {
        int backgroundStillOn = 0;
        foreach (Excel.WorkbookConnection connection in workbook.Connections)
        {
            if (!HasBackgroundQuery(connection)) continue;

            bool oledb = Com.Try(() => connection.OLEDBConnection.BackgroundQuery = false);
            bool odbc = Com.Try(() => connection.ODBCConnection.BackgroundQuery = false);
            if (!oledb && !odbc) backgroundStillOn++;
        }

        (int missingItemsStillKept, int olapCaches) = LimitMissingItems(workbook);

        if (backgroundStillOn > 0)
            reporter.Warning($"{backgroundStillOn} połączeń zostało w trybie odświeżania w tle - zapis może " +
                             "złapać je w połowie odświeżania");
        if (olapCaches > 0)
            reporter.Detail($"{olapCaches} cache tabel przestawnych opiera się na modelu danych - pozycje biorą " +
                            "wprost z modelu, więc kasowanie nieaktualnych ich nie dotyczy");
        if (missingItemsStillKept > 0)
            reporter.Warning($"{missingItemsStillKept} cache tabel przestawnych nie przyjął ustawienia kasowania " +
                             "nieaktualnych pozycji - mogą w nich zostać wartości, których w danych już nie ma");

        ReleaseUnusedComObjects();

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        workbook.RefreshAll();
        App.CalculateUntilAsyncQueriesDone();
        TimeSpan queryElapsed = stopwatch.Elapsed;

        stopwatch.Restart();
        int refreshedCaches = 0;
        int notRefreshed = 0;
        if (refreshPivotCaches)
        {
            foreach (Excel.PivotCache pivotCache in workbook.PivotCaches())
            {
                if (Com.Try(() => pivotCache.Refresh())) refreshedCaches++;
                else notRefreshed++;
            }
            App.CalculateUntilAsyncQueriesDone();

            if (notRefreshed > 0)
                reporter.Warning($"{notRefreshed} cache tabel przestawnych nie dał się odświeżyć - może pamiętać " +
                                 "dane sprzed podziału");
        }

        reporter.Detail($"odświeżenie: zapytania i model {queryElapsed:mm\\:ss}, " +
                        $"cache tabel przestawnych: {(refreshPivotCaches ? $"{refreshedCaches} szt. {stopwatch.Elapsed:mm\\:ss}" : "pominięte")}");
    }

    private static bool HasBackgroundQuery(Excel.WorkbookConnection connection) =>
        Com.Or(() => connection.Type, Excel.XlConnectionType.xlConnectionTypeNOSOURCE)
            is Excel.XlConnectionType.xlConnectionTypeOLEDB
            or Excel.XlConnectionType.xlConnectionTypeODBC
            or Excel.XlConnectionType.xlConnectionTypeDATAFEED;

    private static (int Refused, int Olap) LimitMissingItems(Excel.Workbook workbook)
    {
        int refused = 0;
        int olap = 0;

        foreach (Excel.PivotCache pivotCache in workbook.PivotCaches())
        {
            if (Com.Or(() => pivotCache.OLAP, false))
            {
                olap++;
                continue;
            }

            if (!Com.Try(() => pivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone))
                refused++;
        }

        return (refused, olap);
    }

    public static void CloseWorkbook(Excel.Workbook? workbook, bool save)
    {
        if (workbook is null) return;

        ReleaseUnusedComObjects();
        Com.WhenClosing(() => workbook.Close(SaveChanges: save));
        ReleaseComObject(workbook);
    }

    public static void ReleaseUnusedComObjects()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public static void ReleaseComObject(object? comObject)
    {
        if (comObject is null) return;

        try
        {
            if (Marshal.IsComObject(comObject)) Marshal.FinalReleaseComObject(comObject);
        }
        catch (ArgumentException)
        {
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        var shutdown = Stopwatch.StartNew();

        if (IsResponding())
        {
            var quitting = Stopwatch.StartNew();
            Com.WhenClosing(() =>
            {
                App.DisplayAlerts = false;
                App.Quit();
            });
            Quitting = quitting.Elapsed;
        }
        else if (_excelProcessId > 0)
        {
            ExcelProcessRegistry.TryKill(_excelProcessId);
        }

        ReleaseComObject(App);
        ReleaseUnusedComObjects();

        if (ExcelProcessIsGone()) ExcelProcessRegistry.Unregister(_excelProcessId);
        ExitCode = ExitCodeOf(_excelProcess);
        _excelProcess?.Dispose();
        Shutdown = shutdown.Elapsed;
    }

    private const int GracefulExitMilliseconds = 750;

    private bool ExcelProcessIsGone()
    {
        if (_excelProcessId <= 0) return false;
        try
        {
            using var process = Process.GetProcessById(_excelProcessId);
            if (process.HasExited || process.WaitForExit(GracefulExitMilliseconds)) return true;
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (Exception exception) when (exception is InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            return false;
        }

        return ExcelProcessRegistry.TryKill(_excelProcessId);
    }
}
