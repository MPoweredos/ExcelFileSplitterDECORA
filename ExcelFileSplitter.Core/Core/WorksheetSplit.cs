using System.Diagnostics;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed class WorksheetSplit : ISplitStrategy
{
    private readonly ExcelSession _session;
    private readonly SplitConfig _config;
    private readonly IReadOnlyList<SplitKey> _allSplitKeys;

    private string? _groupedCopyPath;

    public WorksheetSplit(ExcelSession session, SplitConfig config, IReadOnlyList<SplitKey> allSplitKeys)
    {
        _session = session;
        _config = config;
        _allSplitKeys = allSplitKeys;
    }

    public IVerificationResult Verify(
        string outputPath, Recipient recipient, SplitEvidence? evidence, IProgressReporter reporter) =>
        WorksheetVerifier.VerifyRecipientFile(
            outputPath, _config, recipient, _allSplitKeys, evidence, reporter);

    public void PrepareOnce(IProgressReporter reporter)
    {
        string? why = WhyGroupingImpossible();
        if (why is not null)
        {
            reporter.Detail($"grupowanie wierszy pominięte ({why}) - każdy plik zawęzi dane po swojemu");
            return;
        }

        var stopwatch = Stopwatch.StartNew();
        string source = _config.WorkbookToProcess();
        string path = Path.Combine(
            Path.GetTempPath(), $"efs-pogrupowane-{Guid.NewGuid():N}{Path.GetExtension(source)}");

        Excel.Workbook? workbook = null;
        ManualCalculation? calculation = null;
        bool ready = false;
        try
        {
            reporter.Step("Przygotowanie: grupuję wiersze po kluczu podziału (raz dla wszystkich plików)");

            File.Copy(source, path, overwrite: true);
            File.SetAttributes(path, FileAttributes.Normal);

            workbook = _session.OpenWorkbook(path);
            calculation = new ManualCalculation(_session.App);

            DataRange data = FindUnfilteredData(workbook, reporter);

            if (data.LastColumn >= 16_384)
                throw new InvalidOperationException("dane sięgają ostatniej kolumny arkusza - nie ma gdzie postawić klucza");

            FlattenFormulas(workbook, data, reporter);

            SplitKey[] rowKeys = ReadRowKeys(workbook, data);
            var grouping = Stopwatch.StartNew();
            WorksheetTable.GroupRowsByKey(workbook, data, rowKeys, reporter);
            TimeSpan sorted = grouping.Elapsed;

            grouping.Restart();
            workbook.Save();

            _groupedCopyPath = path;
            ready = true;
            reporter.Detail($"pogrupowano {rowKeys.Length:N0} wierszy w {rowKeys.Distinct().Count():N0} klucz(y): " +
                            $"sortowanie {sorted:mm\\:ss}, zapis {grouping.Elapsed:mm\\:ss}, razem {stopwatch.Elapsed:mm\\:ss}");
        }
        catch (Exception exception) when (!Com.SessionIsDead(exception))
        {
            _groupedCopyPath = null;
            reporter.Warning($"nie udało się pogrupować wierszy ({exception.Message}) - " +
                             "podział pójdzie dalej, tyle że wolniej");
        }
        finally
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
            calculation?.Dispose();

            if (!ready) Forget(path);
        }
    }

    public void Dispose()
    {
        if (_groupedCopyPath is null) return;

        Forget(_groupedCopyPath);
        _groupedCopyPath = null;
    }

    private string? WhyGroupingImpossible()
    {
        if (!string.IsNullOrWhiteSpace(_config.Worksheet.Table)) return "dane są tabelą Excela";

        if (!_config.Worksheet.DeleteEntireRow) return "kasowane są same komórki danych, nie całe wiersze";

        return null;
    }

    private void FlattenFormulas(Excel.Workbook workbook, DataRange data, IProgressReporter reporter)
    {
        if (!_config.Worksheet.FormulasToValues) return;

        bool? hasFormulas = WorksheetTable.HasAnyFormula(workbook, data);
        if (hasFormulas == false)
        {
            reporter.Detail("w zakresie danych nie ma formuł - nie ma czego zamieniać");
            return;
        }

        var stopwatch = Stopwatch.StartNew();
        bool done = WorksheetTable.ConvertFormulasToValues(_session.App, workbook, data);

        if (!done)
        {
            reporter.Warning("nie udało się zamienić formuł w danych na wartości - sortowanie będzie wolniejsze");
            return;
        }

        reporter.Detail($"formuły w zakresie danych zamienione na wartości {stopwatch.Elapsed:mm\\:ss}" +
                        (hasFormulas is null ? " (Excel nie potwierdził, czy były - zrobione na wszelki wypadek)" : ""));

        if (DataSheetStaysSomewhere(data.SheetName))
            reporter.Warning($"arkusz '{data.SheetName}' zostaje u odbiorcy, a jego formuły zostały zamienione " +
                             "na wartości - wyłącz Worksheet.FormulasToValues, jeśli mają być żywe");
    }

    private DataRange FindUnfilteredData(Excel.Workbook workbook, IProgressReporter reporter)
    {
        DataRange data = FindData(workbook);

        List<string> cleared = WorksheetTable.ClearFilters(workbook, data.SheetName);
        if (cleared.Count == 0) return data;

        reporter.Detail($"zdjęto filtr ({string.Join(", ", cleared)}) - ukryte nim wiersze też biorą udział " +
                        "w podziale, a pliki odbiorców będą bez filtra");
        return FindData(workbook);
    }

    private DataRange FindData(Excel.Workbook workbook) =>
        WorksheetTable.Find(workbook, _config.Worksheet.Table, _config.Worksheet.Sheet, _config.Worksheet.HeaderRow);

    private bool DataSheetStaysSomewhere(string sheetName) =>
        _config.SheetsToKeep.Concat(_config.Groups.SelectMany(group => group.SheetsToKeep))
            .Any(kept => string.Equals(kept.Trim(), sheetName, StringComparison.OrdinalIgnoreCase));

    private static void Forget(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
        }
    }

    public SplitEvidence? CreateRecipientWorkbook(Recipient recipient, string outputPath, IProgressReporter reporter)
    {
        File.Copy(_groupedCopyPath ?? _config.WorkbookToProcess(), outputPath, overwrite: true);
        File.SetAttributes(outputPath, FileAttributes.Normal);

        Excel.Workbook? workbook = null;
        ManualCalculation? calculation = null;
        try
        {
            var stopwatch = Stopwatch.StartNew();
            workbook = _session.OpenWorkbook(outputPath);
            reporter.Detail($"otwarcie kopii {stopwatch.Elapsed:mm\\:ss}");

            calculation = new ManualCalculation(_session.App);

            NarrowedData narrowed = KeepOnlyRecipientRows(workbook, recipient, reporter);
            PivotRefresh pivots = RebuildCachedViews(workbook, calculation, reporter);

            DataEvidence data = CollectDataEvidence(workbook, recipient, narrowed, reporter);

            stopwatch.Restart();
            SplitExecutor.DeleteSheetsNotKept(workbook, recipient.SheetsToKeep);
            CachedData.ReportRefreshFailures(workbook, pivots, reporter);
            ApplyHardening(workbook, reporter);

            SplitEvidence evidence = CheckFinalState(workbook, data, pivots, reporter);

            workbook.Save();
            reporter.Detail($"czyszczenie, kontrole i zapis {stopwatch.Elapsed:mm\\:ss}");
            return evidence;
        }
        finally
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
            calculation?.Dispose();
        }
    }

    private sealed record NarrowedData(DataRange Range, List<SplitKey> KeptKeys);

    private SplitKey[] ReadRowKeys(Excel.Workbook workbook, DataRange data)
    {
        IReadOnlyList<string> columns = _config.SplitColumnNames;
        var byColumn = new List<string[]>();

        foreach (string column in columns)
            byColumn.Add(WorksheetTable.ReadColumn(workbook, data, WorksheetTable.ColumnIndexOf(data, column)));

        int rows = byColumn.Count == 0 ? 0 : byColumn.Min(values => values.Length);
        var keys = new SplitKey[rows];

        for (int row = 0; row < rows; row++)
        {
            var values = new string[columns.Count];
            for (int column = 0; column < columns.Count; column++) values[column] = byColumn[column][row];
            keys[row] = new SplitKey(values);
        }
        return keys;
    }

    private NarrowedData KeepOnlyRecipientRows(Excel.Workbook workbook, Recipient recipient, IProgressReporter reporter)
    {
        var stopwatch = Stopwatch.StartNew();

        DataRange data = FindUnfilteredData(workbook, reporter);
        SplitKey[] rowKeys = ReadRowKeys(workbook, data);

        var recipientKeys = new HashSet<SplitKey>(recipient.Keys);
        List<RowBlock> blocks = RowBlocks.ToDelete(rowKeys, data.FirstDataRow, recipientKeys.Contains);
        List<SplitKey> keptKeys = rowKeys.Where(recipientKeys.Contains).ToList();

        if (keptKeys.Count == 0 && !_config.Worksheet.AllowEmptyRecipients)
            throw new InvalidOperationException(
                $"w kolumnach podziału ({string.Join(", ", _config.SplitColumnNames)}) nie ma ani jednego wiersza " +
                $"dla tego odbiorcy (sprawdzono {rowKeys.Length:N0} wierszy) - plik byłby pusty");

        if (keptKeys.Count == 0)
            reporter.Warning($"'{recipient.Name}': nie ma dla niego ani jednego wiersza - powstanie pusty plik");

        (long deleted, string how) = DeleteOtherRows(workbook, data, rowKeys, recipientKeys.Contains, blocks, reporter);
        reporter.Detail(
            $"wiersze: zostaje {keptKeys.Count:N0} z {rowKeys.Length:N0}, skasowano {deleted:N0} " +
            $"({blocks.Count} blok(ów), {how}), {stopwatch.Elapsed:mm\\:ss}");

        return new NarrowedData(data with { LastDataRow = data.FirstDataRow + keptKeys.Count - 1 }, keptKeys);
    }

    private (long Deleted, string How) DeleteOtherRows(
        Excel.Workbook workbook, DataRange data, SplitKey[] values, Func<SplitKey, bool> keep, List<RowBlock> blocks,
        IProgressReporter reporter)
    {
        if (WorksheetTable.CanDeleteBySorting(data, _config.Worksheet.DeleteEntireRow, blocks.Count))
        {
            try { return (WorksheetTable.DeleteRowsBySorting(workbook, data, values, keep, reporter), "przez sortowanie"); }
            catch (Exception ex)
            {
                return (WorksheetTable.DeleteRows(_session.App, workbook, data, blocks, _config.Worksheet.DeleteEntireRow),
                        $"blokami, sortowanie zawiodło: {ex.Message}");
            }
        }

        return (WorksheetTable.DeleteRows(_session.App, workbook, data, blocks, _config.Worksheet.DeleteEntireRow), "blokami");
    }

    private sealed record DataEvidence(
        List<SplitKey> KeysInData,
        long KeptRowCount,
        bool DataSheetInFile,
        HashSet<string>? LegalInOtherColumns,
        Dictionary<string, HashSet<string>> ValuesByColumn);

    private DataEvidence CollectDataEvidence(
        Excel.Workbook workbook, Recipient recipient, NarrowedData narrowed, IProgressReporter reporter)
    {
        var stopwatch = Stopwatch.StartNew();

        List<SplitKey> keysInData = narrowed.KeptKeys
            .Where(key => !key.IsIncomplete)
            .Distinct()
            .OrderBy(key => key.Display, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var allowed = new HashSet<string>(
            recipient.Keys.SelectMany(key => key.Values), StringComparer.OrdinalIgnoreCase);
        var forbidden = new HashSet<string>(
            _allSplitKeys.SelectMany(key => key.Values).Where(value => !allowed.Contains(value)),
            StringComparer.OrdinalIgnoreCase);

        var splitColumns = new HashSet<int>(
            _config.SplitColumnNames.Select(column => WorksheetTable.ColumnIndexOf(narrowed.Range, column)));
        HashSet<string>? legalInOtherColumns = FindLegalInOtherColumns(workbook, narrowed.Range, splitColumns, forbidden);
        Dictionary<string, HashSet<string>> valuesByColumn = ReadSlicerColumns(workbook, narrowed.Range, reporter);

        bool dataSheetStays = recipient.SheetsToKeep.Contains(narrowed.Range.SheetName, StringComparer.OrdinalIgnoreCase);

        reporter.Detail($"odczyt danych odbiorcy {stopwatch.Elapsed:mm\\:ss}");
        return new DataEvidence(
            keysInData, narrowed.KeptKeys.Count, dataSheetStays, legalInOtherColumns, valuesByColumn);
    }

    private static SplitEvidence CheckFinalState(
        Excel.Workbook workbook, DataEvidence data, PivotRefresh pivots, IProgressReporter reporter) =>
        new(data.KeysInData,
            data.KeptRowCount,
            data.DataSheetInFile,
            data.LegalInOtherColumns,
            FindStaleSlicerItems(workbook, data.ValuesByColumn, reporter),
            WorksheetTable.FindFormulasWithRefError(workbook))
        {
            StalePivotTables = CachedData.FindStalePivotTables(workbook, pivots),
            PivotRefresh = pivots,
        };

    private static List<string> FindStaleSlicerItems(
        Excel.Workbook workbook, Dictionary<string, HashSet<string>> allowedByColumn, IProgressReporter reporter)
    {
        var staleItems = new List<string>();

        foreach (SlicerInfo slicer in CachedData.ListSlicers(workbook, reporter))
        {
            string field = CachedData.FieldNameOf(slicer.SourceName);
            if (!allowedByColumn.TryGetValue(field, out HashSet<string>? allowedValues)) continue;

            foreach (string caption in SlicerItems.OutsideData(slicer.Captions, allowedValues, SlicerItems.LooksLikeNumberOrDate))
                staleItems.Add($"{slicer.Name}[{field}] -> {caption}");
        }
        return staleItems;
    }

    private static HashSet<string>? FindLegalInOtherColumns(
        Excel.Workbook workbook, DataRange data, HashSet<int> splitColumns, HashSet<string> forbidden)
    {
        if (forbidden.Count == 0) return [];

        try
        {
            var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = data.FirstColumn; column <= data.LastColumn && found.Count < forbidden.Count; column++)
            {
                if (splitColumns.Contains(column)) continue;
                foreach (string value in WorksheetTable.ReadColumn(workbook, data, column))
                    if (forbidden.Contains(value)) found.Add(value);
            }
            return found;
        }
        catch (Exception exception) when (Com.SaysNo(exception) || exception is InvalidOperationException)
        {
            return null;
        }
    }

    private static Dictionary<string, HashSet<string>> ReadSlicerColumns(
        Excel.Workbook workbook, DataRange data, IProgressReporter reporter)
    {
        var valuesByColumn = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (SlicerInfo slicer in CachedData.ListSlicers(workbook, reporter))
        {
            string field = CachedData.FieldNameOf(slicer.SourceName);
            if (field.Length == 0 || valuesByColumn.ContainsKey(field)) continue;

            try
            {
                int column = WorksheetTable.ColumnIndexOf(data, field);
                valuesByColumn[field] = new HashSet<string>(
                    WorksheetTable.ReadColumn(workbook, data, column), StringComparer.OrdinalIgnoreCase);
            }
            catch (InvalidOperationException)
            {
            }
        }
        return valuesByColumn;
    }

    private PivotRefresh RebuildCachedViews(Excel.Workbook workbook, ManualCalculation calculation, IProgressReporter reporter)
    {
        var stopwatch = Stopwatch.StartNew();

        calculation.CalculateOnce();

        PivotRefresh pivots = CachedData.PurgePivotCaches(workbook, reporter);
        int charts = CachedData.RefreshChartCaches(workbook, reporter);

        string caches = pivots.Failures.Count == 0
            ? $"{pivots.Refreshed}"
            : $"{pivots.Refreshed} z {pivots.Refreshed + pivots.Failures.Count}";
        reporter.Detail(
            $"odświeżenie: {caches} cache tabel przestawnych, {charts} wykres(ów), {stopwatch.Elapsed:mm\\:ss}");
        return pivots;
    }

    private void ApplyHardening(Excel.Workbook workbook, IProgressReporter reporter)
    {
        HardeningOptions hardening = _config.Hardening;

        if (hardening.RemoveSplitColumnSlicers)
        {
            foreach (string column in _config.SplitColumnNames)
            {
                int removed = CachedData.RemoveSlicersBoundToColumn(workbook, column, reporter);
                if (removed > 0) reporter.Detail($"usunięto {removed} fragmentator(ów) na '{column}'");
            }
        }

        if (hardening.DisableRefreshOnOpen) CachedData.DisableRefreshOnOpen(workbook, reporter);

        if (hardening.ProtectStructure)
        {
            if (!Com.Try(() => workbook.Protect(Structure: true, Windows: false)))
                reporter.Warning("nie udało się zablokować struktury skoroszytu");
        }
    }
}
