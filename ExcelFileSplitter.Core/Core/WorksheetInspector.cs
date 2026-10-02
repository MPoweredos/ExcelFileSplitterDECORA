using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public static class WorksheetInspector
{
    public static WorkbookInfo ReadWorkbookInfo(
        ExcelSession session, SplitConfig config, IProgressReporter reporter, bool readSplitValues = true)
    {
        return session.WithWorkbook(
            config.WorkbookToProcess(), readOnly: true,
            workbook => ReadOpenWorkbook(session, workbook, config, reporter, readSplitValues));
    }

    public static WorkbookInfo ReadOpenWorkbook(
        ExcelSession session, Excel.Workbook workbook, SplitConfig config, IProgressReporter reporter,
        bool readSplitValues = true, WorkbookInfo? reuse = null)
    {
        List<SheetInfo>? sheets = reuse is null ? Interop.Sheets.List(workbook) : null;
        List<string> sheetNames = reuse?.Sheets ?? sheets!.Select(sheet => sheet.Name).ToList();
        List<string> visibleSheets =
            reuse?.VisibleSheets ?? sheets!.Where(sheet => sheet.IsVisible).Select(sheet => sheet.Name).ToList();
        List<SheetDependency> dependencies = reuse?.Dependencies ?? SheetDependencies.FindDependencies(workbook);

        var dataProblems = new List<string>();
        var dataTables = new List<ModelTableInfo>();
        var splitKeys = new List<SplitKey>();

        DataRange? data = FindData(workbook, config, dataProblems);

        if (data is not null)
        {
            FilterClearing filters;
            try
            {
                filters = WorksheetTable.TryClearFilters(workbook, data.SheetName);
            }
            catch (Exception exception) when (Com.SaysNo(exception) && !Com.SessionIsDead(exception))
            {
                dataProblems.Add($"nie udało się sprawdzić, czy dane mają włączony filtr ({exception.Message.Trim()}) - " +
                                 "przy włączonym filtrze podział zostawiłby w plikach wiersze innych odbiorców");
                filters = new FilterClearing([], []);
            }

            if (filters.Remaining.Count > 0) dataProblems.Add(filters.Problem());
            if (filters.Cleared.Count > 0)
            {
                reporter.Detail($"dane mają włączony filtr ({string.Join(", ", filters.Cleared)}) - zdjęty na czas " +
                                "odczytu i podziału; ukryte nim wiersze też biorą udział w podziale, a pliki " +
                                "odbiorców będą bez filtra. Plik źródłowy zostaje bez zmian.");
                data = FindData(workbook, config, dataProblems);
            }
        }

        if (data is not null)
        {
            dataTables.Add(Describe(data));

            if (data.EmptyRowsAfterData > 0)
                reporter.Detail(
                    $"arkusz '{data.SheetName}' uważa za używane wiersze do {data.SheetLastRow:N0}, " +
                    $"a dane kończą się na {data.LastDataRow:N0} - {data.EmptyRowsAfterData:N0} pustych wierszy " +
                    "za tabelą zostało pominiętych");

            dataProblems.AddRange(FindDataProblems(session, workbook, config, data));
            if (readSplitValues) splitKeys = ReadSplitKeys(workbook, config, data, dataProblems, reporter);
        }

        return new WorkbookInfo(
            sheetNames,
            visibleSheets,
            dataTables,
            [],
            splitKeys,
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            dependencies,
            [],
            dataProblems,
            data?.SheetName);
    }

    private static DataRange? FindData(Excel.Workbook workbook, SplitConfig config, List<string> dataProblems)
    {
        try
        {
            return WorksheetTable.Find(workbook, config.Worksheet.Table, config.Worksheet.Sheet, config.Worksheet.HeaderRow);
        }
        catch (Exception ex)
        {
            dataProblems.Add(ex.Message);
            return null;
        }
    }

    private static ModelTableInfo Describe(DataRange data) => new(
        data.TableName ?? $"{data.SheetName} (wiersz nagłówków: {data.HeaderRow})",
        data.Headers.Select(header => new ModelColumn(header, 0, false)).ToList(),
        data.RowCount);

    private static List<string> FindDataProblems(
        ExcelSession session, Excel.Workbook workbook, SplitConfig config, DataRange data)
    {
        var problems = new List<string>();

        foreach (string column in config.SplitColumnNames)
        {
            try { WorksheetTable.ColumnIndexOf(data, column); }
            catch (Exception ex) { problems.Add(ex.Message); }
        }

        if (data.RowCount == 0)
            problems.Add($"{data.Describe()} nie ma żadnych wierszy danych");

        if (config.Worksheet.DeleteEntireRow)
        {
            long? beside = WorksheetTable.CountCellsBesideData(session.App, workbook, data);
            if (beside is null)
                problems.Add(
                    "nie udało się sprawdzić, czy obok danych, w tych samych wierszach, stoi cokolwiek jeszcze - " +
                    "a przy Worksheet.DeleteEntireRow = true kasowanie całych wierszy zabrałoby to razem z danymi. " +
                    "Sprawdź arkusz ręcznie albo ustaw Worksheet.DeleteEntireRow = false.");
            else if (beside > 0)
                problems.Add(
                    $"obok danych, w tych samych wierszach, stoi {beside:N0} niepustych komórek - kasowanie całych " +
                    "wierszy zabrałoby je razem z danymi. Ustaw Worksheet.DeleteEntireRow = false (kasowane będą wtedy " +
                    "wyłącznie komórki tabeli) albo przenieś tę zawartość gdzie indziej.");
        }

        return problems;
    }

    private static List<SplitKey> ReadSplitKeys(
        Excel.Workbook workbook, SplitConfig config, DataRange data, List<string> dataProblems,
        IProgressReporter reporter)
    {
        IReadOnlyList<string> columns = config.SplitColumnNames;
        var byColumn = new List<string[]>();

        foreach (string column in columns)
        {
            try
            {
                byColumn.Add(WorksheetTable.ReadColumn(workbook, data, WorksheetTable.ColumnIndexOf(data, column)));
            }
            catch (Exception ex)
            {
                reporter.Warning($"nie udało się odczytać wartości kolumny '{column}': {ex.Message}");
                return [];
            }
        }

        if (byColumn.Count == 0) return [];

        int rows = byColumn.Min(values => values.Length);
        var keys = new List<SplitKey>();
        var seen = new HashSet<SplitKey>();
        long incomplete = 0;

        for (int row = 0; row < rows; row++)
        {
            var values = new string[columns.Count];
            for (int column = 0; column < columns.Count; column++) values[column] = byColumn[column][row];

            var key = new SplitKey(values);
            if (key.IsIncomplete)
            {
                incomplete++;
                continue;
            }
            if (seen.Add(key)) keys.Add(key);
        }

        if (incomplete > 0)
            dataProblems.Add(
                $"{incomplete:N0} wierszy nie ma kompletu wartości w kolumnach podziału " +
                $"({string.Join(", ", columns)}) - taki wiersz nie pasuje do żadnego odbiorcy i zniknie ze " +
                "WSZYSTKICH plików. Uzupełnij te komórki albo świadomie pomiń te kontrole.");

        keys.Sort((first, second) => string.Compare(first.Display, second.Display, StringComparison.OrdinalIgnoreCase));
        return keys;
    }
}
