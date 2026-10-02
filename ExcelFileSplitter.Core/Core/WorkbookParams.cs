using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public static class WorkbookParams
{
    public const string ParameterColumnHeader = "Parameter";
    public const string ValueColumnHeader = "Value";

    public static bool TrySetParameter(Excel.Workbook workbook, string tableName, string parameter, string value)
    {
        try { SetParameter(workbook, tableName, parameter, value); return true; }
        catch (InvalidOperationException) { return false; }
    }

    public static void SetParameter(Excel.Workbook workbook, string tableName, string parameter, string value)
    {
        Excel.ListObject paramsTable = WorksheetTable.FindTable(workbook, tableName)
            ?? throw new InvalidOperationException(
                $"Nie znaleziono tabeli parametrów '{tableName}' w skoroszycie. " +
                "Zobacz docs/workbook-setup.md - trzeba ją raz założyć.");

        Excel.Range dataRows = paramsTable.DataBodyRange
            ?? throw new InvalidOperationException($"Tabela '{tableName}' nie ma wierszy danych.");

        int parameterColumn = FindColumnIndex(paramsTable, ParameterColumnHeader);
        int valueColumn = FindColumnIndex(paramsTable, ValueColumnHeader);

        for (int row = 1; row <= dataRows.Rows.Count; row++)
        {
            object? parameterCell = (dataRows.Cells[row, parameterColumn] as Excel.Range)?.Value2;
            if (string.Equals(Convert.ToString(parameterCell), parameter, StringComparison.OrdinalIgnoreCase))
            {
                var valueCell = (Excel.Range)dataRows.Cells[row, valueColumn];
                valueCell.NumberFormat = "@";
                valueCell.Value2 = value;
                return;
            }
        }

        throw new InvalidOperationException(
            $"W tabeli '{tableName}' nie ma wiersza o parametrze '{parameter}'. Dopisz go w skoroszycie.");
    }

    private static int FindColumnIndex(Excel.ListObject table, string header)
    {
        int index = 1;
        foreach (Excel.ListColumn column in table.ListColumns)
        {
            if (string.Equals(column.Name?.Trim(), header, StringComparison.OrdinalIgnoreCase)) return index;
            index++;
        }
        throw new InvalidOperationException(
            $"Tabela '{table.Name}' nie ma kolumny '{header}'. Wymagane kolumny: {ParameterColumnHeader}, {ValueColumnHeader}.");
    }
}
