using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed class ModelQuery : IDisposable
{
    private readonly dynamic _adoConnection;

    public ModelQuery(Excel.Workbook workbook)
    {
        Excel.Model model = workbook.Model;

        _ = model.ModelTables.Count;

        _adoConnection = model.DataModelConnection.ModelConnection.ADOConnection;
        Com.Try(() => _adoConnection.CommandTimeout = 0);
    }

    public void StreamRows(string query, Action<string[]> onHeader, Action<object?[]> onRow, int chunkSize = 20_000)
    {
        dynamic recordset = _adoConnection.Execute(query);
        try
        {
            int columnCount = (int)recordset.Fields.Count;
            var columnNames = new string[columnCount];
            for (int i = 0; i < columnCount; i++)
                columnNames[i] = ExtractColumnName((string)recordset.Fields[i].Name);
            onHeader(columnNames);

            while (!(bool)recordset.EOF)
            {
                object data = recordset.GetRows(chunkSize);
                if (data is not Array chunk) break;

                int firstColumn = chunk.GetLowerBound(0);
                int firstRow = chunk.GetLowerBound(1);
                int rowCount = chunk.GetLength(1);
                if (rowCount == 0) break;

                for (int row = 0; row < rowCount; row++)
                {
                    var values = new object?[columnCount];
                    for (int column = 0; column < columnCount; column++)
                    {
                        object? value = chunk.GetValue(firstColumn + column, firstRow + row);
                        values[column] = value is DBNull ? null : value;
                    }
                    onRow(values);
                }
            }
        }
        finally
        {
            Com.Try(() => recordset.Close());
            ExcelSession.ReleaseComObject(recordset);
        }
    }

    public (string[] Columns, List<object?[]> Rows) ReadAllRows(string query)
    {
        string[] columns = [];
        var rows = new List<object?[]>();
        StreamRows(query, header => columns = header, row => rows.Add(row));
        return (columns, rows);
    }

    public List<string> ReadDistinctValues(string table, string column)
    {
        string columnRef = ColumnRef(table, column);
        string dax = $"""
            EVALUATE
            SELECTCOLUMNS(
                FILTER(VALUES({columnRef}), NOT ISBLANK({columnRef})),
                "Value", {columnRef}
            )
            ORDER BY [Value]
            """;

        var (_, rows) = ReadAllRows(dax);
        return rows
            .Select(row => Convert.ToString(row[0], System.Globalization.CultureInfo.CurrentCulture) ?? "")
            .Where(value => value.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    public long CountTableRows(string table)
    {
        var (_, rows) = ReadAllRows($"EVALUATE ROW(\"n\", COUNTROWS({TableRef(table)}))");
        if (rows.Count == 0 || rows[0][0] is null) return 0;
        return Convert.ToInt64(rows[0][0]);
    }

    public long CountOrphanKeys(string factTable, string factColumn, string dimensionTable, string dimensionColumn)
    {
        var (_, rows) = ReadAllRows(
            $"EVALUATE ROW(\"n\", COUNTROWS(EXCEPT(VALUES({ColumnRef(factTable, factColumn)}), VALUES({ColumnRef(dimensionTable, dimensionColumn)}))))");
        if (rows.Count == 0 || rows[0][0] is null) return 0;
        return Convert.ToInt64(rows[0][0]);
    }

    public HashSet<string> FindValuesPresentInColumns(
        IEnumerable<(string Table, string Column)> columns, IEnumerable<string> candidates)
    {
        var candidateList = candidates.Where(value => !string.IsNullOrEmpty(value)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var columnList = columns.ToList();
        var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (candidateList.Count == 0 || columnList.Count == 0) return present;

        string candidateRows = UnionOf(candidateList.Select(value => $"ROW(\"v\", {TextLiteral(value)})"));
        string allColumnValues = UnionOf(columnList.Select(column =>
        {
            string columnRef = ColumnRef(column.Table, column.Column);
            return $"SELECTCOLUMNS(VALUES({columnRef}), \"v\", {columnRef} & \"\")";
        }));

        var (_, rows) = ReadAllRows($"EVALUATE INTERSECT({candidateRows}, {allColumnValues})");
        foreach (var row in rows)
            if (row[0] is string value) present.Add(value);
        return present;

        static string UnionOf(IEnumerable<string> parts)
        {
            var list = parts.ToList();
            return list.Count == 1 ? list[0] : "UNION(" + string.Join(", ", list) + ")";
        }
    }

    public HashSet<string> FindCalculatedColumns()
    {
        var calculated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var (columns, rows) = ReadAllRows("SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY");
            int objectType = IndexOf(columns, "OBJECT_TYPE"), table = IndexOf(columns, "TABLE"), objectName = IndexOf(columns, "OBJECT");
            int referencedType = IndexOf(columns, "REFERENCED_OBJECT_TYPE");
            int referencedTable = IndexOf(columns, "REFERENCED_TABLE"), referencedObject = IndexOf(columns, "REFERENCED_OBJECT");
            foreach (var row in rows)
            {
                if (FieldEquals(row, objectType, "CALC_COLUMN")) calculated.Add(ColumnKey(AsText(row[table]), AsText(row[objectName])));
                if (FieldEquals(row, referencedType, "CALC_COLUMN")) calculated.Add(ColumnKey(AsText(row[referencedTable]), AsText(row[referencedObject])));
            }
        }
        catch (Exception exception) when (Com.SaysNo(exception))
        {
        }

        try
        {
            var (columns, rows) = ReadAllRows("SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS");
            int dimension = IndexOf(columns, "DIMENSION_NAME"), attribute = IndexOf(columns, "ATTRIBUTE_NAME");
            int columnId = IndexOf(columns, "COLUMN_ID"), columnType = IndexOf(columns, "COLUMN_TYPE");
            foreach (var row in rows)
                if (FieldEquals(row, columnType, "BASIC_DATA") && AsText(row[columnId]).StartsWith("Calculated Column", StringComparison.OrdinalIgnoreCase))
                    calculated.Add(ColumnKey(AsText(row[dimension]), AsText(row[attribute])));
        }
        catch (Exception exception) when (Com.SaysNo(exception))
        {
        }

        return calculated;

        static int IndexOf(string[] columns, string name) =>
            Array.FindIndex(columns, column => string.Equals(column, name, StringComparison.OrdinalIgnoreCase));
        static bool FieldEquals(object?[] row, int index, string expected) =>
            index >= 0 && string.Equals(AsText(row[index]), expected, StringComparison.OrdinalIgnoreCase);
        static string AsText(object? value) => Convert.ToString(value) ?? "";
    }

    public static string ColumnKey(string table, string column) => table + "|" + column;

    public static string TableRef(string table) => "'" + table.Replace("'", "''") + "'";

    public static string ColumnRef(string table, string column) => TableRef(table) + "[" + column.Replace("]", "]]") + "]";

    public static string TextLiteral(string text) => "\"" + text.Replace("\"", "\"\"") + "\"";

    private static string ExtractColumnName(string header)
    {
        int open = header.LastIndexOf('[');
        int close = header.LastIndexOf(']');
        if (open >= 0 && close == header.Length - 1 && close > open)
            return header.Substring(open + 1, close - open - 1);
        return header;
    }

    public void Dispose() => ExcelSession.ReleaseComObject(_adoConnection);
}
