using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed record DataRange(
    string SheetName,
    string? TableName,
    int HeaderRow,
    int FirstDataRow,
    int LastDataRow,
    int FirstColumn,
    int LastColumn,
    IReadOnlyList<string> Headers)
{
    public int SheetLastRow { get; init; }

    public int EmptyRowsAfterData => SheetLastRow > LastDataRow ? SheetLastRow - LastDataRow : 0;

    public long RowCount => LastDataRow < FirstDataRow ? 0 : LastDataRow - FirstDataRow + 1L;

    public string Describe() => TableName is not null
        ? $"tabela '{TableName}' na arkuszu '{SheetName}'"
        : $"arkusz '{SheetName}', nagłówki w wierszu {HeaderRow}";
}

public sealed record FilterClearing(List<string> Cleared, List<string> Remaining)
{
    public string Problem() =>
        $"nie udało się zdjąć filtra: {string.Join(", ", Remaining)}. Przy włączonym filtrze Excel " +
        "kasuje i sortuje tylko widoczne wiersze, więc w plikach zostałyby dane innych odbiorców. " +
        "Zdejmij filtr w pliku źródłowym (albo ochronę arkusza, jeśli jest) i dopiero wtedy uruchom podział.";
}

public static class WorksheetTable
{
    private const int BlocksPerDeleteCall = 200;

    private const int BlocksWhereSortingWins = 340;

    public static DataRange Find(Excel.Workbook workbook, string tableName, string sheetName, int headerRow)
    {
        if (!string.IsNullOrWhiteSpace(tableName))
        {
            Excel.ListObject table = FindTable(workbook, tableName)
                ?? throw new InvalidOperationException(
                    $"Nie znaleziono tabeli Excela '{tableName}'. Dostępne tabele: {string.Join(", ", ListTableNames(workbook))}");
            return FromTable(table);
        }

        if (string.IsNullOrWhiteSpace(sheetName))
            throw new InvalidOperationException("Konfiguracja nie wskazuje danych - ustaw Worksheet.Table albo Worksheet.Sheet.");

        return FromSheet(workbook, sheetName, headerRow);
    }

    public static Excel.Worksheet GetSheet(Excel.Workbook workbook, string sheetName)
    {
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return sheet;

        throw new InvalidOperationException($"Nie znaleziono arkusza '{sheetName}' w skoroszycie.");
    }

    public static List<string> ClearFilters(Excel.Workbook workbook, string sheetName)
    {
        FilterClearing result = TryClearFilters(workbook, sheetName);
        if (result.Remaining.Count > 0) throw new InvalidOperationException(result.Problem());
        return result.Cleared;
    }

    public static FilterClearing TryClearFilters(Excel.Workbook workbook, string sheetName) =>
        TryClearFilters(GetSheet(workbook, sheetName));

    public static FilterClearing TryClearAllFilters(Excel.Workbook workbook)
    {
        var cleared = new List<string>();
        var remaining = new List<string>();
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            FilterClearing one = TryClearFilters(sheet);
            cleared.AddRange(one.Cleared);
            remaining.AddRange(one.Remaining);
        }
        return new FilterClearing(cleared, remaining);
    }

    private static FilterClearing TryClearFilters(Excel.Worksheet sheet)
    {
        List<string> active = ActiveFilters(sheet, whenUnknown: false);
        if (active.Count == 0) return new FilterClearing([], []);

        if (Com.Or(() => sheet.FilterMode, false)) Com.Try(() => sheet.ShowAllData());
        foreach (Excel.ListObject table in sheet.ListObjects)
            if (IsFiltered(table, whenUnknown: false)) Com.Try(() => table.AutoFilter.ShowAllData());

        List<string> remaining = ActiveFilters(sheet, whenUnknown: true);
        return new FilterClearing(active.Except(remaining).ToList(), remaining);
    }

    private static List<string> ActiveFilters(Excel.Worksheet sheet, bool whenUnknown)
    {
        var active = new List<string>();
        if (Com.Or(() => sheet.FilterMode, whenUnknown)) active.Add($"arkusz '{sheet.Name}'");

        foreach (Excel.ListObject table in sheet.ListObjects)
            if (IsFiltered(table, whenUnknown)) active.Add($"tabela '{table.Name}'");
        return active;
    }

    private static bool IsFiltered(Excel.ListObject table, bool whenUnknown) =>
        Com.Or(() => table.ShowAutoFilter && table.AutoFilter is { FilterMode: true }, whenUnknown);

    private static void EnsureNoHiddenRows(Excel.Worksheet sheet, DataRange data)
    {
        if (data.RowCount == 0) return;

        Excel.Range rows = sheet.Range[sheet.Cells[data.FirstDataRow, 1], sheet.Cells[data.LastDataRow, 1]].EntireRow;

        object? unknown = new object();
        object? hidden = Com.Or<object?>(() => rows.Hidden, unknown);
        if (hidden is false) return;

        throw new InvalidOperationException(ReferenceEquals(hidden, unknown)
            ? "Excel nie potwierdził, że w danych nie ma ukrytych wierszy"
            : "w danych są ukryte wiersze, a sortowanie nie przesuwa ukrytych");
    }

    public static int ColumnIndexOf(DataRange data, string header)
    {
        for (int i = 0; i < data.Headers.Count; i++)
            if (string.Equals(data.Headers[i], header, StringComparison.OrdinalIgnoreCase))
                return data.FirstColumn + i;

        throw new InvalidOperationException(
            $"W danych ({data.Describe()}) nie ma kolumny '{header}'. Kolumny: {string.Join(", ", data.Headers)}");
    }

    public static string[] ReadColumn(Excel.Workbook workbook, DataRange data, int column)
    {
        if (data.RowCount == 0) return [];

        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        Excel.Range range = sheet.Range[sheet.Cells[data.FirstDataRow, column], sheet.Cells[data.LastDataRow, column]];
        object? cells = range.Value2;

        if (cells is not object[,] grid) return [AsText(cells)];

        int firstRow = grid.GetLowerBound(0);
        int firstColumn = grid.GetLowerBound(1);
        var values = new string[grid.GetLength(0)];
        for (int i = 0; i < values.Length; i++) values[i] = AsText(grid[firstRow + i, firstColumn]);
        return values;
    }

    public static long DeleteRows(
        Excel.Application app, Excel.Workbook workbook, DataRange data, IReadOnlyList<RowBlock> blocks, bool entireRow)
    {
        if (blocks.Count == 0) return 0;

        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);

        for (int last = blocks.Count - 1; last >= 0; last -= BlocksPerDeleteCall)
        {
            int first = Math.Max(0, last - BlocksPerDeleteCall + 1);
            Excel.Range? batch = null;

            for (int i = last; i >= first; i--)
            {
                Excel.Range rows = RangeOf(sheet, data, blocks[i], entireRow);
                batch = batch is null ? rows : app.Union(batch, rows);
            }

            batch!.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        return RowBlocks.TotalRows(blocks);
    }

    public static long? CountCellsBesideData(Excel.Application app, Excel.Workbook workbook, DataRange data)
    {
        if (data.RowCount == 0) return 0;

        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        Excel.Range used = sheet.UsedRange;
        int usedFirstColumn = used.Column;
        int usedLastColumn = usedFirstColumn + used.Columns.Count - 1;

        long count = 0;
        if (usedFirstColumn < data.FirstColumn)
        {
            long? left = CountNonEmpty(app, sheet, data.FirstDataRow, usedFirstColumn, data.LastDataRow, data.FirstColumn - 1);
            if (left is null) return null;
            count += left.Value;
        }
        if (usedLastColumn > data.LastColumn)
        {
            long? right = CountNonEmpty(app, sheet, data.FirstDataRow, data.LastColumn + 1, data.LastDataRow, usedLastColumn);
            if (right is null) return null;
            count += right.Value;
        }
        return count;
    }

    public static long DeleteRowsBySorting<T>(
        Excel.Workbook workbook, DataRange data, IReadOnlyList<T> values, Func<T, bool> keep,
        IProgressReporter reporter)
    {
        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        EnsureNoHiddenRows(sheet, data);

        int helperColumn = data.LastColumn + 1;
        int firstRow = data.FirstDataRow;
        int lastRow = data.LastDataRow;

        var sortKeys = new object[values.Count, 1];
        long keptCount = 0;
        for (int i = 0; i < values.Count; i++)
        {
            bool stays = keep(values[i]);
            if (stays) keptCount++;
            sortKeys[i, 0] = (stays ? 0d : 1_000_000_000d) + i;
        }

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        Excel.Range helperRange = sheet.Range[sheet.Cells[firstRow, helperColumn], sheet.Cells[lastRow, helperColumn]];
        helperRange.Value2 = sortKeys;
        TimeSpan keyElapsed = stopwatch.Elapsed;

        try
        {
            stopwatch.Restart();
            Excel.Range sortRange = sheet.Range[sheet.Cells[firstRow, data.FirstColumn], sheet.Cells[lastRow, helperColumn]];
            sortRange.Sort(
                Key1: sheet.Cells[firstRow, helperColumn],
                Order1: Excel.XlSortOrder.xlAscending,
                Header: Excel.XlYesNoGuess.xlNo,
                Orientation: Excel.XlSortOrientation.xlSortColumns,
                MatchCase: false,
                DataOption1: Excel.XlSortDataOption.xlSortNormal);
            TimeSpan sortElapsed = stopwatch.Elapsed;

            stopwatch.Restart();
            long toDelete = values.Count - keptCount;
            if (toDelete > 0)
            {
                int firstDeletedRow = firstRow + (int)keptCount;
                sheet.Range[sheet.Cells[firstDeletedRow, 1], sheet.Cells[lastRow, 1]].EntireRow
                    .Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }

            reporter.Detail($"sortowanie: klucz {keyElapsed:mm\\:ss}, sort {sortElapsed:mm\\:ss}, " +
                            $"kasowanie {stopwatch.Elapsed:mm\\:ss}");
            return toDelete;
        }
        finally
        {
            bool cleared = Com.Try(() =>
                sheet.Range[sheet.Cells[firstRow, helperColumn], sheet.Cells[lastRow, helperColumn]].ClearContents());

            if (!cleared)
                reporter.Warning($"nie udało się wyczyścić kolumny pomocniczej użytej do sortowania " +
                                 $"(kolumna {helperColumn} arkusza '{data.SheetName}') - sprawdź plik wynikowy");
        }
    }

    public static bool? HasAnyFormula(Excel.Workbook workbook, DataRange data)
    {
        if (data.RowCount == 0) return false;

        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        Excel.Range range = sheet.Range[
            sheet.Cells[data.FirstDataRow, data.FirstColumn], sheet.Cells[data.LastDataRow, data.LastColumn]];

        object? mixed = new object();
        object? answer = Com.Or<object?>(() => range.HasFormula, mixed);

        if (ReferenceEquals(answer, mixed)) return null;
        if (answer is bool only) return only;
        return true;
    }

    public static bool ConvertFormulasToValues(Excel.Application app, Excel.Workbook workbook, DataRange data)
    {
        if (data.RowCount == 0) return true;

        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        Excel.Range range = sheet.Range[
            sheet.Cells[data.FirstDataRow, data.FirstColumn], sheet.Cells[data.LastDataRow, data.LastColumn]];

        bool done = Com.Try(() =>
        {
            range.Copy();
            range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
        });

        Com.WhenClosing(() => app.CutCopyMode = (Excel.XlCutCopyMode)0);
        return done;
    }

    public static void GroupRowsByKey<T>(
        Excel.Workbook workbook, DataRange data, IReadOnlyList<T> rowKeys, IProgressReporter reporter)
        where T : notnull
    {
        Excel.Worksheet sheet = GetSheet(workbook, data.SheetName);
        EnsureNoHiddenRows(sheet, data);

        int helperColumn = data.LastColumn + 1;
        int firstRow = data.FirstDataRow;
        int lastRow = data.LastDataRow;

        var rankOf = new Dictionary<T, int>();
        var sortKeys = new object[rowKeys.Count, 1];
        for (int i = 0; i < rowKeys.Count; i++)
        {
            if (!rankOf.TryGetValue(rowKeys[i], out int rank))
            {
                rank = rankOf.Count;
                rankOf[rowKeys[i]] = rank;
            }
            sortKeys[i, 0] = (rank * 1_000_000_000d) + i;
        }

        Excel.Range helperRange = sheet.Range[sheet.Cells[firstRow, helperColumn], sheet.Cells[lastRow, helperColumn]];
        helperRange.Value2 = sortKeys;

        try
        {
            Excel.Range sortRange = sheet.Range[sheet.Cells[firstRow, data.FirstColumn], sheet.Cells[lastRow, helperColumn]];
            sortRange.Sort(
                Key1: sheet.Cells[firstRow, helperColumn],
                Order1: Excel.XlSortOrder.xlAscending,
                Header: Excel.XlYesNoGuess.xlNo,
                Orientation: Excel.XlSortOrientation.xlSortColumns,
                MatchCase: false,
                DataOption1: Excel.XlSortDataOption.xlSortNormal);
        }
        finally
        {
            bool cleared = Com.Try(() =>
                sheet.Range[sheet.Cells[firstRow, helperColumn], sheet.Cells[lastRow, helperColumn]].ClearContents());

            if (!cleared)
                reporter.Warning($"nie udało się wyczyścić kolumny pomocniczej użytej do grupowania " +
                                 $"(kolumna {helperColumn} arkusza '{data.SheetName}') - sprawdź pliki wynikowe");
        }
    }

    public static bool CanDeleteBySorting(DataRange data, bool entireRow, int blockCount) =>
        data.TableName is null
        && entireRow
        && blockCount >= BlocksWhereSortingWins
        && data.LastColumn < 16_384;

    public static List<string> FindFormulasWithRefError(Excel.Workbook workbook, int maxReported = 20)
    {
        var broken = new List<string>();

        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            var errorCells = Com.Or<Excel.Range?>(
                () => sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors),
                null);
            if (errorCells is null) continue;

            foreach (Excel.Range area in errorCells.Areas)
            {
                object? areaFormulas = Com.Or<object?>(() => area.Formula, null);
                if (areaFormulas is null) continue;

                IEnumerable<object?> cells = areaFormulas is object[,] grid
                    ? grid.Cast<object?>()
                    : new object?[] { areaFormulas };

                foreach (object? cell in cells)
                {
                    if (cell is not string formula || !formula.Contains("#REF!", StringComparison.Ordinal)) continue;
                    broken.Add($"{sheet.Name}: {Truncate(formula, 60)}");
                    if (broken.Count >= maxReported) return broken;
                }
            }
        }
        return broken;
    }

    public static List<string> ListTableNames(Excel.Workbook workbook)
    {
        var names = new List<string>();
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            foreach (Excel.ListObject table in sheet.ListObjects)
                names.Add(table.Name);
        return names;
    }

    public static List<DataRange> ListTables(Excel.Workbook workbook)
    {
        var tables = new List<DataRange>();
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            foreach (Excel.ListObject table in sheet.ListObjects)
            {
                try
                {
                    tables.Add(FromTable(table));
                }
                catch (InvalidOperationException)
                {
                }
            }
        return tables;
    }

    public static DataRange? TryFind(Excel.Workbook workbook, string tableName, string sheetName, int headerRow)
    {
        try { return Find(workbook, tableName, sheetName, headerRow); }
        catch (InvalidOperationException) { return null; }
        catch (Exception exception) when (Com.SaysNo(exception)) { return null; }
    }

    private static long? CountNonEmpty(
        Excel.Application app, Excel.Worksheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        Com.Or<long?>(() =>
        {
            Excel.Range range = sheet.Range[sheet.Cells[firstRow, firstColumn], sheet.Cells[lastRow, lastColumn]];
            return (long)app.WorksheetFunction.CountA(range);
        }, null);

    private static Excel.Range RangeOf(Excel.Worksheet sheet, DataRange data, RowBlock block, bool entireRow) =>
        entireRow
            ? sheet.Range[sheet.Cells[block.First, 1], sheet.Cells[block.Last, 1]].EntireRow
            : sheet.Range[sheet.Cells[block.First, data.FirstColumn], sheet.Cells[block.Last, data.LastColumn]];

    private static DataRange FromTable(Excel.ListObject table)
    {
        var sheet = (Excel.Worksheet)table.Parent;

        Excel.Range headerRange = table.HeaderRowRange
            ?? throw new InvalidOperationException(
                $"Tabela '{table.Name}' nie ma wiersza nagłówków - włącz go w Excelu albo wskaż dane przez Worksheet.Sheet.");

        int firstColumn = headerRange.Column;
        int headerRow = headerRange.Row;

        Excel.Range? body = table.DataBodyRange;
        int lastDataRow = body is null ? headerRow : body.Row + body.Rows.Count - 1;

        return new DataRange(
            sheet.Name, table.Name, headerRow, headerRow + 1, lastDataRow,
            firstColumn, firstColumn + headerRange.Columns.Count - 1, ReadHeaders(headerRange));
    }

    private static DataRange FromSheet(Excel.Workbook workbook, string sheetName, int headerRow)
    {
        Excel.Worksheet sheet = GetSheet(workbook, sheetName);

        Excel.Range used = sheet.UsedRange;
        int usedFirstRow = used.Row;
        int usedFirstColumn = used.Column;
        int usedLastRow = usedFirstRow + used.Rows.Count - 1;
        int usedLastColumn = usedFirstColumn + used.Columns.Count - 1;

        if (headerRow < usedFirstRow || headerRow > usedLastRow)
            throw new InvalidOperationException(
                $"Na arkuszu '{sheet.Name}' wiersz nagłówków {headerRow} jest poza obszarem z danymi " +
                $"(arkusz używa wierszy {usedFirstRow}-{usedLastRow}). Popraw Worksheet.HeaderRow.");

        Excel.Range headerRange = sheet.Range[sheet.Cells[headerRow, usedFirstColumn], sheet.Cells[headerRow, usedLastColumn]];
        List<string> headers = ReadHeaders(headerRange);

        while (headers.Count > 0 && headers[^1].Length == 0) headers.RemoveAt(headers.Count - 1);
        if (headers.Count == 0)
            throw new InvalidOperationException($"W wierszu {headerRow} arkusza '{sheet.Name}' nie ma żadnych nagłówków.");

        int lastColumn = usedFirstColumn + headers.Count - 1;
        int lastDataRow = LastRowWithData(sheet, usedFirstColumn, lastColumn, usedLastRow) ?? usedLastRow;

        if (lastDataRow < headerRow) lastDataRow = headerRow;

        return new DataRange(
            sheet.Name, null, headerRow, headerRow + 1, lastDataRow,
            usedFirstColumn, lastColumn, headers)
        {
            SheetLastRow = usedLastRow,
        };
    }

    private static int? LastRowWithData(Excel.Worksheet sheet, int firstColumn, int lastColumn, int lastRow)
    {
        return Com.Or<int?>(() =>
        {
            Excel.Range columns = sheet.Range[sheet.Cells[1, firstColumn], sheet.Cells[lastRow, lastColumn]];

            Excel.Range found = columns.Find(
                What: "*",
                LookIn: Excel.XlFindLookIn.xlFormulas,
                LookAt: Excel.XlLookAt.xlPart,
                SearchOrder: Excel.XlSearchOrder.xlByRows,
                SearchDirection: Excel.XlSearchDirection.xlPrevious);

            return found?.Row;
        }, null);
    }

    private static List<string> ReadHeaders(Excel.Range headerRange)
    {
        object? cells = Com.Or<object?>(() => headerRange.Value2, null);

        if (cells is not object[,] grid) return [AsText(cells)];

        int firstRow = grid.GetLowerBound(0);
        int firstColumn = grid.GetLowerBound(1);
        var headers = new List<string>(grid.GetLength(1));
        for (int i = 0; i < grid.GetLength(1); i++) headers.Add(AsText(grid[firstRow, firstColumn + i]));
        return headers;
    }

    internal static Excel.ListObject? FindTable(Excel.Workbook workbook, string name)
    {
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            foreach (Excel.ListObject table in sheet.ListObjects)
                if (string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase))
                    return table;
        return null;
    }

    private static string Truncate(string text, int maxLength) => text.Length <= maxLength ? text : text[..maxLength] + "...";

    private static string AsText(object? value) => value switch
    {
        null => "",
        string text => text.Trim(),
        double number => number.ToString(CultureInfo.InvariantCulture),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "",
    };
}
