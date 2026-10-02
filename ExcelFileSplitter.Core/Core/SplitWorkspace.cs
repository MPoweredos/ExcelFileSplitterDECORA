using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed class SplitWorkspace : IDisposable
{
    private ExcelSession? _session;
    private Excel.Workbook? _workbook;
    private bool _readsSheetData;
    private bool _disposed;

    public string? WorkbookPath { get; private set; }

    public WorkbookOutline? Outline { get; private set; }

    private ExcelSession Session
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return _session ??= new ExcelSession();
        }
    }

    private Excel.Workbook OpenedWorkbook
    {
        get
        {
            if (_workbook is not null) return _workbook;
            if (WorkbookPath is null)
                throw new InvalidOperationException("Skoroszyt nie jest otwarty - najpierw wskaż plik do podziału.");

            _workbook = OpenSource(WorkbookPath, _readsSheetData, reporter: null);
            return _workbook;
        }
    }

    private Excel.Workbook OpenSource(string path, bool clearFilters, IProgressReporter? reporter)
    {
        Excel.Workbook workbook = Session.OpenWorkbook(path, readOnly: true);
        if (!clearFilters) return workbook;

        FilterClearing filters;
        try
        {
            filters = WorksheetTable.TryClearAllFilters(workbook);
        }
        catch (Exception exception) when (Com.SaysNo(exception) && !Com.SessionIsDead(exception))
        {
            reporter?.Warning($"nie udało się sprawdzić filtrów ({exception.Message.Trim()}) - kontrola gotowości " +
                              "sprawdzi je jeszcze raz na arkuszu z danymi");
            return workbook;
        }
        catch
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
            throw;
        }

        if (reporter is not null && filters.Cleared.Count > 0)
            reporter.Detail($"zdjęto filtr ({string.Join(", ", filters.Cleared)}) - tylko w otwartym pliku, " +
                            "źródło zostaje bez zmian; liczby wierszy i wartości obejmują też wiersze, które filtr ukrywał");
        if (reporter is not null && filters.Remaining.Count > 0)
            reporter.Warning($"nie udało się zdjąć filtra ({string.Join(", ", filters.Remaining)}) - jeśli są tam " +
                             "dane do podziału, kontrola gotowości nie pozwoli zacząć; zdejmij filtr w pliku źródłowym");

        return workbook;
    }

    private T Do<T>(Func<T> work)
    {
        try
        {
            return work();
        }
        catch (Exception exception) when (Com.SessionIsDead(exception))
        {
            ReplaceSession();
            throw new InvalidOperationException(
                $"Excel przestał odpowiadać i jego proces został zamknięty ({exception.Message.Trim()}). " +
                "Kolejna próba uruchomi Excela od nowa - powtórz ostatni krok.", exception);
        }
    }

    private void ReplaceSession()
    {
        _workbook = null;

        ExcelSession? previous = _session;
        _session = null;
        previous?.Dispose();
    }

    private T RetryOnFreshExcel<T>(Func<T> work, IProgressReporter reporter)
    {
        try
        {
            return Do(work);
        }
        catch (Exception exception) when (Com.WorthRetryingOnFreshExcel(exception))
        {
            reporter.Warning($"Excel nie poradził sobie z modelem danych ({exception.Message.Trim()}) - " +
                             "uruchamiam go od nowa i próbuję jeszcze raz");

            ReplaceSession();
            return Do(work);
        }
    }

    public WorkbookOutline OpenWorkbook(string path, IProgressReporter reporter, bool findDataCandidates = true) =>
        Do(() => OpenWorkbookCore(path, reporter, findDataCandidates));

    private WorkbookOutline OpenWorkbookCore(string path, IProgressReporter reporter, bool findDataCandidates)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Nie ma pliku '{path}'.", path);

        Forget();

        string fullPath = Path.GetFullPath(path);
        reporter.Step($"Otwieram {Path.GetFileName(fullPath)}");
        Session.ReportStartupProblems(reporter);
        _workbook = OpenSource(fullPath, findDataCandidates, reporter);
        _readsSheetData = findDataCandidates;
        WorkbookPath = fullPath;

        var sheets = Interop.Sheets.List(_workbook);
        var candidates = new List<DataCandidate>();

        if (findDataCandidates)
        {
            candidates.AddRange(WorksheetTable.ListTables(_workbook).Select(DataCandidate.From));

            var sheetsWithTables = candidates.Select(candidate => candidate.SheetName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (SheetInfo sheet in sheets)
            {
                if (sheetsWithTables.Contains(sheet.Name)) continue;
                DataRange? range = WorksheetTable.TryFind(_workbook, tableName: "", sheet.Name, headerRow: 1);
                if (range is not null) candidates.Add(DataCandidate.From(range));
            }
        }

        var dependencies = SheetDependencies.FindDependencies(_workbook);
        reporter.Detail($"arkusze: {sheets.Count}, kandydaci na dane: {candidates.Count}");

        Outline = new WorkbookOutline(fullPath, sheets, candidates, dependencies);
        return Outline;
    }

    public DataCandidate ReadData(WorksheetOptions source) =>
        Do(() => DataCandidate.From(WorksheetTable.Find(OpenedWorkbook, source.Table, source.Sheet, source.HeaderRow)));

    public List<SplitKeyCount> ReadSplitKeys(WorksheetOptions source, IReadOnlyList<string> columns) =>
        Do(() => ReadSplitKeysCore(source, columns));

    private List<SplitKeyCount> ReadSplitKeysCore(WorksheetOptions source, IReadOnlyList<string> columns)
    {
        if (columns.Count == 0) return [];

        DataRange data = WorksheetTable.Find(OpenedWorkbook, source.Table, source.Sheet, source.HeaderRow);

        var byColumn = new List<string[]>();
        foreach (string column in columns)
            byColumn.Add(WorksheetTable.ReadColumn(OpenedWorkbook, data, WorksheetTable.ColumnIndexOf(data, column)));

        int rows = byColumn.Min(values => values.Length);
        var keys = new List<SplitKey>(rows);

        for (int row = 0; row < rows; row++)
        {
            var values = new string[columns.Count];
            for (int column = 0; column < columns.Count; column++) values[column] = byColumn[column][row];
            keys.Add(new SplitKey(values));
        }

        return SplitKeyCount.Tally(keys);
    }

    public WorkbookInfo Inspect(
        SplitConfig config, IProgressReporter reporter, bool readSplitValues = true, WorkbookInfo? reuse = null) =>
        config.Mode == SplitMode.Worksheet
            ? Do(() => WorksheetInspector.ReadOpenWorkbook(
                Session, OpenedWorkbook, config, reporter, readSplitValues, reuse))
            : RetryOnFreshExcel(
                () => WorkbookInspector.ReadOpenWorkbook(Session, OpenedWorkbook, config, reporter, readSplitValues),
                reporter);

    public void CloseWorkbook()
    {
        if (_workbook is null) return;

        try
        {
            ExcelSession.CloseWorkbook(_workbook, save: false);
            _workbook = null;
        }
        catch (Exception exception) when (Com.SessionIsDead(exception))
        {
            ReplaceSession();
        }
    }

    private void Forget()
    {
        CloseWorkbook();
        WorkbookPath = null;
        Outline = null;
        _readsSheetData = false;
    }

    public void Reset()
    {
        try
        {
            CloseWorkbook();
        }
        finally
        {
            WorkbookPath = null;
            Outline = null;
            _readsSheetData = false;
            ReplaceSession();
        }
    }

    public WorkbookOutline OpenForProcessing(SplitConfig config, IProgressReporter reporter)
    {
        string path = config.WorkbookToProcess();
        bool findDataCandidates = config.Mode == SplitMode.Worksheet;

        return config.Mode == SplitMode.Worksheet
            ? Do(() => OpenWorkbookCore(path, reporter, findDataCandidates))
            : RetryOnFreshExcel(() => OpenWorkbookCore(path, reporter, findDataCandidates), reporter);
    }

    public List<SplitKey> ReadModelValues(string table, string column, IProgressReporter reporter) =>
        RetryOnFreshExcel(() =>
        {
            using var model = new ModelQuery(OpenedWorkbook);
            return model.ReadDistinctValues(table, column).Select(value => SplitKey.Of(value)).ToList();
        }, reporter);

    public string Prepare(SplitConfig config, IProgressReporter reporter) => Do(() =>
    {
        CloseWorkbook();
        return WorkbookPreparer.PrepareWorkbook(Session, config, reporter);
    });

    public CacheResult BuildCache(SplitConfig config, IProgressReporter reporter)
    {
        try
        {
            return Do(() =>
            {
                CloseWorkbook();
                return ModelCacheBuilder.BuildCache(Session, config, reporter);
            });
        }
        finally
        {
            if (_session is not null)
            {
                reporter.Detail("wymieniam proces Excela po odświeżeniu modelu - kolejny etap zaczyna od czystego");
                ReplaceSession();
            }
        }
    }

    public List<PassRun> InspectAll(SplitConfig config, IProgressReporter reporter, bool readSplitValues = true) =>
        Do(() => SplitRun.PlanAll(
            config, (forPass, reuse) => Inspect(forPass, reporter, readSplitValues, reuse), reporter));

    public List<SplitOutcome> Split(
        IReadOnlyList<PassRun> runs,
        IProgressReporter reporter,
        CancellationToken cancellation = default,
        IReadOnlySet<string>? onlyThese = null) => Do(() =>
    {
        CloseWorkbook();
        try
        {
            return SplitRun.Execute(Session, runs, reporter, cancellation, onlyThese);
        }
        finally
        {
            int killed = ExcelProcessRegistry.KillOwn(except: _session?.ProcessId);
            if (killed > 0)
                reporter.Warning($"po podziale zostało {killed} proces(ów) Excela, które powinny były się zamknąć - " +
                                 "zamknąłem je na siłę");
        }
    });

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Forget();
        _session?.Dispose();
        _session = null;
    }
}
