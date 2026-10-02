using System.Diagnostics;
using System.Xml;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed class PowerPivotSplit : ISplitStrategy
{
    private readonly SplitConfig _config;
    private readonly IReadOnlyList<string> _allSplitColumnValues;
    private readonly IReadOnlyList<ModelTableInfo> _referenceTables;
    private List<CollapsedPivot> _collapsedPivots = [];

    public PowerPivotSplit(
        SplitConfig config, IReadOnlyList<SplitKey> allSplitKeys, IReadOnlyList<ModelTableInfo> referenceTables)
    {
        _config = config;
        _allSplitColumnValues = allSplitKeys.Select(key => key.Values[0]).ToList();
        _referenceTables = referenceTables;
    }

    public void PrepareOnce(IProgressReporter reporter)
    {
        string path = _config.WorkbookToProcess();
        try
        {
            _collapsedPivots = CollapsedPivotLayout.Find(path, _config.PowerPivot.NarrowedTables());
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException
                                              or UnauthorizedAccessException)
        {
            _collapsedPivots = [];
            reporter.Warning($"nie udało się odczytać układu tabel przestawnych z {Path.GetFileName(path)} " +
                             $"({exception.Message.Trim()}) - zwinięte tabele nie zostaną rozwinięte na czas " +
                             "odświeżenia; czy nie pamiętają pozycji innych odbiorców, sprawdzi weryfikacja");
        }

        foreach (CollapsedPivot pivot in _collapsedPivots)
            reporter.Detail($"{pivot.Describe()} - w każdym pliku, do którego trafia jej arkusz, zostanie " +
                            "rozwinięta na czas odświeżenia i zwinięta z powrotem");
    }

    public void Dispose() { }

    private static ExcelSession Fresh(string what, IProgressReporter reporter)
    {
        var session = new ExcelSession();
        reporter.Detail($"{what}start Excela {session.Startup.Text()}");
        return session;
    }

    private static void Close(string what, ExcelSession session, Excel.Workbook? workbook, IProgressReporter reporter)
    {
        var closing = Stopwatch.StartNew();
        ExcelSession.CloseWorkbook(workbook, save: false);
        session.Dispose();

        reporter.Detail($"{what}zamknięcie Excela {closing.Elapsed.Text()} ({session.DescribeShutdown()})");
        if (session.Crashed)
            reporter.Warning($"{what}Excel zakończył się awarią (kod {session.ExitCodeText()}). Jeśli w tle pojawił " +
                             "się wtedy pusty Excel, to Windows uruchomił go ponownie po tej awarii - można go zamknąć");
    }

    public IVerificationResult Verify(
        string outputPath, Recipient recipient, SplitEvidence? evidence, IProgressReporter reporter)
    {
        ExcelSession session = Fresh("weryfikacja: ", reporter);
        try
        {
            return SplitVerifier.VerifyRecipientFile(
                session, outputPath, _config, recipient, _allSplitColumnValues, _referenceTables, reporter);
        }
        finally
        {
            Close("weryfikacja: ", session, workbook: null, reporter);
        }
    }

    public SplitEvidence? CreateRecipientWorkbook(Recipient recipient, string outputPath, IProgressReporter reporter)
    {
        DateTime refreshStarted;
        try
        {
            refreshStarted = CopyAndProcess(recipient, outputPath, reporter);
        }
        catch (Exception exception) when (Com.SessionIsDead(exception))
        {
            reporter.Warning($"Excel przestał działać w trakcie tworzenia pliku ({exception.Message.Trim()}) - " +
                             "zaczynam ten plik od nowa, na nowym Excelu");
            refreshStarted = CopyAndProcess(recipient, outputPath, reporter);
        }

        CheckQueriesRefreshedInFile(outputPath, refreshStarted, reporter);
        return null;
    }

    private DateTime CopyAndProcess(Recipient recipient, string outputPath, IProgressReporter reporter)
    {
        var copying = Stopwatch.StartNew();
        File.Copy(_config.WorkbookToProcess(), outputPath, overwrite: true);
        File.SetAttributes(outputPath, FileAttributes.Normal);
        reporter.Detail($"kopia skoroszytu {copying.Elapsed.Text()}");

        try
        {
            return ProcessInExcel(recipient, outputPath, reporter);
        }
        catch
        {
            DeleteUnprocessedCopy(outputPath, reporter);
            throw;
        }
    }

    private static void DeleteUnprocessedCopy(string outputPath, IProgressReporter reporter)
    {
        try
        {
            File.Delete(outputPath);
            reporter.Detail($"skasowano nieprzetworzoną kopię {Path.GetFileName(outputPath)} - miała dane " +
                            "wszystkich odbiorców");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            reporter.Warning($"nie udało się skasować nieprzetworzonej kopii {Path.GetFileName(outputPath)} " +
                             $"({exception.Message.Trim()}) - ma dane wszystkich odbiorców");
        }
    }

    private DateTime ProcessInExcel(Recipient recipient, string outputPath, IProgressReporter reporter)
    {
        ExcelSession session = Fresh("", reporter);

        Excel.Workbook? workbook = null;
        DateTime refreshStarted;
        try
        {
            var opening = Stopwatch.StartNew();
            workbook = session.OpenWorkbook(outputPath);
            reporter.Detail($"otwarcie skoroszytu {opening.Elapsed.Text()}");

            var parameters = Stopwatch.StartNew();

            WorkbookParams.SetParameter(
                workbook, _config.PowerPivot.ParamsTable, _config.PowerPivot.UseCacheParamName,
                _config.PowerPivot.UseCache ? "1" : "0");
            WorkbookParams.SetParameter(
                workbook, _config.PowerPivot.ParamsTable, _config.PowerPivot.SplitParamName,
                string.Join("\n", recipient.Keys.Select(key => key.Values[0])));
            if (!string.IsNullOrWhiteSpace(_config.PowerPivot.CacheFolder))
            {
                string cacheFolder = Path.GetFullPath(_config.PowerPivot.CacheFolder);
                if (!cacheFolder.EndsWith(Path.DirectorySeparatorChar)) cacheFolder += Path.DirectorySeparatorChar;
                WorkbookParams.SetParameter(
                    workbook, _config.PowerPivot.ParamsTable, _config.PowerPivot.CachePathParamName, cacheFolder);
            }

            reporter.Detail($"parametry podziału {parameters.Elapsed.Text()}");

            ExcelSession.ReleaseUnusedComObjects();
            List<CollapsedPivot> expanded = CollapsedPivots.ExpandBeforeRefresh(
                workbook,
                _collapsedPivots.Where(pivot =>
                    recipient.SheetsToKeep.Contains(pivot.Table.Sheet, StringComparer.OrdinalIgnoreCase)),
                reporter);

            refreshStarted = DateTime.Now;

            var stopwatch = Stopwatch.StartNew();
            session.RefreshAllAndWait(workbook, _config.PowerPivot.RefreshPivotCaches, reporter);

            var notRefreshed = ModelSchema.FindTablesNotRefreshed(
                refreshStarted, ModelSchema.ReadRefreshDates(workbook),
                _config.PowerPivot.Queries.Select(query => query.Name));
            if (notRefreshed.Count > 0)
                throw new InvalidOperationException(
                    $"nie odświeżyły się tabele: {string.Join(", ", notRefreshed)} - zostałyby w nich dane sprzed podziału. " +
                    $"Odświeżenie zaczęło się {refreshStarted:yyyy-MM-dd HH\\:mm\\:ss}. " +
                    "Sprawdź zapytania w przygotowanym skoroszycie (edytor Power Query).");

            CollapsedPivots.CollapseBack(workbook, expanded, reporter);

            int refreshedCharts = CachedData.RefreshChartCaches(workbook, reporter);
            reporter.Detail($"odświeżenie {stopwatch.Elapsed.Text()} ({refreshedCharts} wykres(ów))");

            stopwatch.Restart();
            SplitExecutor.DeleteSheetsNotKept(workbook, recipient.SheetsToKeep);
            ApplyHardening(workbook, reporter);
            ExcelSession.ReleaseUnusedComObjects();
            workbook.Save();
            reporter.Detail($"czyszczenie i zapis {stopwatch.Elapsed.Text()}");
        }
        finally
        {
            Close("", session, workbook, reporter);
        }

        return refreshStarted;
    }

    private void CheckQueriesRefreshedInFile(string outputPath, DateTime refreshStarted, IProgressReporter reporter)
    {
        Dictionary<string, QueryFill>? fills = QueryRefreshLog.Read(outputPath);
        List<string> queries = _config.PowerPivot.Queries.Select(query => query.Name).ToList();
        if (fills is null)
        {
            reporter.Detail("w pliku nie ma zapisu odświeżenia zapytań Power Query - nie da się sprawdzić go w pliku");
            return;
        }

        reporter.Detail("wiersze po odświeżeniu: " + string.Join(", ", queries.Select(query =>
            fills.TryGetValue(query, out QueryFill? fill) && fill.Rows is long rows ? $"{query} {rows:N0}" : $"{query} ?")));

        List<QueryFill> stale = QueryRefreshLog.NotRefreshed(fills, queries, refreshStarted.ToUniversalTime());
        if (stale.Count == 0) return;

        string details = string.Join("; ", stale.Select(fill =>
            $"'{fill.Query}' - " + (fill.LastFilledUtc is DateTime lastFilled
                ? $"ostatnio {lastFilled.ToLocalTime():yyyy-MM-dd HH\\:mm}" + (fill.Rows is long rows ? $", {rows:N0} wierszy" : "")
                : "brak zapisu odświeżenia")));
        throw new InvalidOperationException(
            $"zapytania nie odświeżyły się przy podziale: {details}. W modelu zostały ich dane sprzed podziału, " +
            "czyli dane wszystkich odbiorców. Otwórz przygotowany skoroszyt, odśwież te zapytania i sprawdź, jaki " +
            "błąd zgłasza Power Query (np. brak dostępu do plików źródłowych albo do cache CSV) i czy są " +
            "włączone w „Odśwież wszystko”.");
    }

    private void ApplyHardening(Excel.Workbook workbook, IProgressReporter reporter)
    {
        HardeningOptions hardening = _config.Hardening;

        if (hardening.RemoveSplitColumnSlicers)
        {
            int removed = CachedData.RemoveSlicersBoundToColumn(workbook, _config.SplitColumn, reporter);
            if (removed > 0) reporter.Detail($"usunięto {removed} fragmentator(ów) na '{_config.SplitColumn}'");
        }

        if (hardening.DisableRefreshOnOpen) CachedData.DisableRefreshOnOpen(workbook, reporter);

        if (hardening.StubQueries)
            PowerQuery.ReplaceAllFormulasWithError(
                workbook, "Zapytanie zostało usunięte przy podziale pliku.", reporter);

        if (hardening.DeleteQueries)
            PowerQuery.DeleteAllQueries(workbook, reporter);

        if (hardening.ProtectStructure)
        {
            if (!Com.Try(() => workbook.Protect(Structure: true, Windows: false)))
                reporter.Warning("nie udało się zablokować struktury skoroszytu");
        }
    }
}
