using System.Diagnostics;
using ExcelFileSplitter.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public static class CollapsedPivots
{
    public static List<CollapsedPivot> ExpandBeforeRefresh(
        Excel.Workbook workbook, IEnumerable<CollapsedPivot> layouts, IProgressReporter reporter)
    {
        var expanded = new List<CollapsedPivot>();

        foreach (CollapsedPivot layout in layouts)
        {
            Excel.Worksheet? sheet = SheetOf(workbook, layout.Table.Sheet);
            Excel.PivotTable? table = sheet is null ? null : TableOn(sheet, layout.Table.Name);
            if (sheet is null || table is null)
            {
                reporter.Warning($"nie znalazłem w otwartym pliku tabeli przestawnej {layout.Table} - nie zostanie " +
                                 "rozwinięta na czas odświeżenia; czy nie pamięta pozycji innych odbiorców, " +
                                 "sprawdzi weryfikacja");
                continue;
            }

            string? blocker = WhatIsBelow(sheet, table);
            if (blocker is not null)
            {
                reporter.Warning(
                    $"{layout.Describe()}. Nie rozwijam jej na czas odświeżenia, bo {blocker}. Jeśli weryfikacja " +
                    "zatrzyma plik, rozwiń tę tabelę w pliku źródłowym na wszystkich poziomach („Rozwiń całe pole”, " +
                    "ang. „Expand Entire Field”) albo przenieś to, co jest pod nią, i przygotuj plik od nowa");
                continue;
            }

            var stopwatch = Stopwatch.StartNew();
            int rowsBefore = RowsOf(table);
            var refused = new List<string>();
            foreach (CollapsedField field in layout.Fields)
                if (!Expand(table, field)) refused.Add(field.Caption);
            int rowsAfter = RowsOf(table);

            expanded.Add(layout);
            if (refused.Count == 0 && rowsAfter > rowsBefore)
                reporter.Detail($"rozwinięto zwiniętą tabelę przestawną {layout.Table} na czas odświeżenia, żeby " +
                                "odświeżenie przebudowało też ukryte poziomy z danych odbiorcy (wiersze tabeli: " +
                                $"{rowsBefore} → {rowsAfter}, {stopwatch.Elapsed.Text()})");
            else
                reporter.Warning($"rozwinięcie tabeli przestawnej {layout.Table} na czas odświeżenia nie zadziałało (" +
                                 (refused.Count > 0 ? $"Excel nie przyjął pól: {string.Join(", ", refused)}; " : "") +
                                 $"wiersze tabeli: {rowsBefore} → {rowsAfter}) - może zachować pozycje z pliku " +
                                 "źródłowego; sprawdzi to weryfikacja");
        }

        return expanded;
    }

    public static void CollapseBack(
        Excel.Workbook workbook, IReadOnlyList<CollapsedPivot> expanded, IProgressReporter reporter)
    {
        foreach (CollapsedPivot layout in expanded)
        {
            Excel.Worksheet? sheet = SheetOf(workbook, layout.Table.Sheet);
            Excel.PivotTable? table = sheet is null ? null : TableOn(sheet, layout.Table.Name);
            if (table is null)
            {
                reporter.Warning($"po odświeżeniu nie znalazłem tabeli przestawnej {layout.Table} - zostaje " +
                                 "rozwinięta; pokazuje dane odbiorcy, różni się tylko wyglądem");
                continue;
            }

            int rowsBefore = RowsOf(table);
            var failed = new List<string>();
            foreach (CollapsedField field in layout.Fields.Reverse())
                if (!Collapse(table, field)) failed.Add(field.Caption);
            int rowsAfter = RowsOf(table);

            if (failed.Count == 0)
                reporter.Detail($"zwinięto z powrotem tabelę przestawną {layout.Table} " +
                                $"(wiersze tabeli: {rowsBefore} → {rowsAfter})");
            else
                reporter.Warning($"nie udało się zwinąć z powrotem tabeli przestawnej {layout.Table} " +
                                 $"(pola: {string.Join(", ", failed)}; wiersze tabeli: {rowsBefore} → {rowsAfter}) - " +
                                 "zostaje w pliku rozwinięta; pokazuje dane odbiorcy, różni się tylko wyglądem");
        }
    }

    private static bool Expand(Excel.PivotTable table, CollapsedField field)
    {
        Excel.PivotField? pivotField = FieldOf(table, field.Name);
        return pivotField is not null && Com.Try(() => pivotField.DrilledDown = true);
    }

    private static bool Collapse(Excel.PivotTable table, CollapsedField field)
    {
        Excel.PivotField? pivotField = FieldOf(table, field.Name);
        if (pivotField is null) return false;
        if (field.AllCollapsed) return Com.Try(() => pivotField.DrilledDown = false);

        foreach (string name in field.CollapsedItems)
        {
            Excel.PivotItem? item = ItemOf(pivotField, name);
            if (item is not null) Com.Try(() => item.DrilledDown = false);
        }

        return true;
    }

    private static Excel.PivotField? FieldOf(Excel.PivotTable table, string name)
    {
        var field = Com.Or<Excel.PivotField?>(() => (Excel.PivotField)table.PivotFields(name), null);
        if (field is not null) return field;

        var rowFields = Com.Or<Excel.PivotFields?>(() => (Excel.PivotFields)table.RowFields, null);
        int count = rowFields is null ? 0 : Com.Or(() => rowFields.Count, 0);
        for (int i = 1; i <= count; i++)
        {
            int index = i;
            var candidate = Com.Or<Excel.PivotField?>(() => (Excel.PivotField)rowFields!.Item(index), null);
            if (candidate is not null
                && string.Equals(Com.Or(() => candidate.Name, ""), name, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }

        return null;
    }

    private static Excel.PivotItem? ItemOf(Excel.PivotField field, string name) =>
        Com.Or<Excel.PivotItem?>(() => (Excel.PivotItem)field.PivotItems(name), null);

    private static int RowsOf(Excel.PivotTable table) => Com.Or(() => table.TableRange1.Rows.Count, -1);

    private static Excel.Worksheet? SheetOf(Excel.Workbook workbook, string name) =>
        Com.Or<Excel.Worksheet?>(() => (Excel.Worksheet)workbook.Worksheets[name], null);

    private static Excel.PivotTable? TableOn(Excel.Worksheet sheet, string name) =>
        Com.Or<Excel.PivotTable?>(() => (Excel.PivotTable)sheet.PivotTables(name), null);

    private static string? WhatIsBelow(Excel.Worksheet sheet, Excel.PivotTable table)
    {
        try
        {
            Excel.Range area = table.TableRange2;
            int lastRow = area.Row + area.Rows.Count - 1;
            int firstColumn = area.Column;
            int lastColumn = firstColumn + area.Columns.Count - 1;
            int sheetRows = sheet.Rows.Count;
            if (lastRow >= sheetRows) return "kończy się na ostatnim wierszu arkusza";

            Excel.Range below = sheet.Range[sheet.Cells[lastRow + 1, firstColumn], sheet.Cells[sheetRows, lastColumn]];
            Excel.Range? found = below.Find(
                What: "*",
                LookIn: Excel.XlFindLookIn.xlFormulas,
                LookAt: Excel.XlLookAt.xlPart,
                SearchOrder: Excel.XlSearchOrder.xlByRows,
                SearchDirection: Excel.XlSearchDirection.xlNext);

            return found is null
                ? null
                : $"pod nią, w komórce {found.Address[false, false]}, są inne dane, które rozwinięcie mogłoby nadpisać";
        }
        catch (Exception exception) when (Com.SaysNo(exception) && !Com.SessionIsDead(exception))
        {
            return $"nie da się sprawdzić, co jest pod nią ({ExcelMessage.FirstSentence(exception.Message)})";
        }
    }
}
