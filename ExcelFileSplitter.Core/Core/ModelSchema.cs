using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed record ModelColumn(string Name, int DataType, bool IsCalculated);

public sealed record ModelTableInfo(string Name, List<ModelColumn> Columns, long RowCount);

public sealed record StaleTable(string Name, DateTime RefreshedAt)
{
    public override string ToString() => $"{Name} (data odświeżenia {RefreshedAt:yyyy-MM-dd HH\\:mm\\:ss})";
}

public static class ModelSchema
{
    public static List<ModelTableInfo> ReadTables(Excel.Workbook workbook, ModelQuery model, bool countRows)
    {
        var calculatedColumns = model.FindCalculatedColumns();
        var tables = new List<ModelTableInfo>();

        foreach (Excel.ModelTable modelTable in workbook.Model.ModelTables)
        {
            var columns = new List<ModelColumn>();
            foreach (Excel.ModelTableColumn modelColumn in modelTable.ModelTableColumns)
            {
                int dataType = Com.Or(() => Convert.ToInt32(modelColumn.DataType), 0);
                columns.Add(new ModelColumn(
                    modelColumn.Name, dataType, calculatedColumns.Contains(ModelQuery.ColumnKey(modelTable.Name, modelColumn.Name))));
            }

            long rowCount = countRows ? Com.Or(() => model.CountTableRows(modelTable.Name), 0L) : 0L;
            tables.Add(new ModelTableInfo(modelTable.Name, columns, rowCount));
        }
        return tables;
    }

    public static List<string> FindDifferences(IReadOnlyList<ModelTableInfo> expected, IReadOnlyList<ModelTableInfo> actual)
    {
        var differences = new List<string>();
        var actualByName = actual.ToDictionary(table => table.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var expectedTable in expected)
        {
            if (!actualByName.TryGetValue(expectedTable.Name, out var actualTable))
            {
                differences.Add($"brak tabeli '{expectedTable.Name}'");
                continue;
            }

            var actualColumns = actualTable.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
            var expectedNames = new HashSet<string>(expectedTable.Columns.Select(column => column.Name), StringComparer.OrdinalIgnoreCase);

            foreach (var expectedColumn in expectedTable.Columns)
            {
                if (!actualColumns.TryGetValue(expectedColumn.Name, out var actualColumn))
                    differences.Add($"{expectedTable.Name}: brak kolumny '{expectedColumn.Name}'");
                else if (actualColumn.DataType != expectedColumn.DataType)
                    differences.Add($"{expectedTable.Name}[{expectedColumn.Name}]: typ '{DescribeDataType(actualColumn.DataType)}' " +
                                    $"zamiast '{DescribeDataType(expectedColumn.DataType)}'");
            }

            foreach (var actualColumn in actualTable.Columns)
                if (!expectedNames.Contains(actualColumn.Name))
                    differences.Add($"{expectedTable.Name}: dodatkowa kolumna '{actualColumn.Name}'");
        }

        var expectedTableNames = new HashSet<string>(expected.Select(table => table.Name), StringComparer.OrdinalIgnoreCase);
        foreach (var actualTable in actual)
            if (!expectedTableNames.Contains(actualTable.Name))
                differences.Add($"dodatkowa tabela '{actualTable.Name}'");

        return differences;
    }

    public static Dictionary<string, DateTime?> ReadRefreshDates(Excel.Workbook workbook)
    {
        var refreshDates = new Dictionary<string, DateTime?>(StringComparer.OrdinalIgnoreCase);
        foreach (Excel.ModelTable modelTable in workbook.Model.ModelTables)
        {
            refreshDates[modelTable.Name] =
                Com.Or<DateTime?>(() => Convert.ToDateTime(((dynamic)modelTable).RefreshDate), null);
        }
        return refreshDates;
    }

    public static List<StaleTable> FindTablesNotRefreshed(
        DateTime refreshStarted, Dictionary<string, DateTime?> afterRefresh, IEnumerable<string> tables)
    {
        DateTime threshold = RefreshClock.Threshold(refreshStarted);

        var stale = new List<StaleTable>();
        foreach (string name in tables)
        {
            if (!afterRefresh.TryGetValue(name, out DateTime? refreshed) || refreshed is null) continue;

            if (refreshed.Value < threshold) stale.Add(new StaleTable(name, refreshed.Value));
        }
        return stale;
    }

    public static string DescribeDataType(int adoType) => adoType switch
    {
        130 or 129 or 200 or 201 or 202 or 203 => "tekst",
        7 or 133 or 134 or 135 => "data",
        2 or 3 or 16 or 17 or 18 or 19 or 20 or 21 => "liczba całkowita",
        4 or 5 or 14 or 131 => "liczba",
        6 => "waluta",
        11 => "logiczny",
        _ => adoType.ToString(),
    };
}
