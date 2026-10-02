using System.Diagnostics;
using System.Globalization;
using System.Text;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public enum ColumnKind { Unknown, Text, Int, Number, DateTime, Logical }

public sealed record TableSchema(
    string TableName, string CsvPath, string[] ColumnNames, string[] ColumnTypes, long RowCount, string[] SkippedCalculatedColumns);

public sealed record CacheResult(IReadOnlyList<TableSchema> DumpedTables, bool SelfCheckRan, List<string> SelfCheckProblems)
{
    public bool SelfCheckPassed => SelfCheckProblems.Count == 0;
}

public static class ModelCacheBuilder
{
    public static CacheResult BuildCache(ExcelSession session, SplitConfig config, IProgressReporter reporter)
    {
        string cacheFolder = Path.GetFullPath(config.PowerPivot.CacheFolder);
        Directory.CreateDirectory(cacheFolder);
        EnsureCachedQueriesHaveTables(config);

        Excel.Workbook? workbook = null;
        try
        {
            reporter.Step($"Otwieram {config.WorkbookToProcess()}");
            workbook = session.OpenWorkbook(config.WorkbookToProcess());

            bool hasParamsTable = SetParametersForFullDump(workbook, config, cacheFolder);
            if (!hasParamsTable)
                reporter.Warning($"brak tabeli '{config.PowerPivot.ParamsTable}' - odświeżam skoroszyt w stanie domyślnym.");

            var stopwatch = Stopwatch.StartNew();
            reporter.Step("Odświeżam z plików źródłowych (to ta wolna część)...");
            session.RefreshAllAndWait(workbook, config.PowerPivot.RefreshPivotCaches, reporter);
            reporter.Detail($"odświeżenie ze źródeł: {stopwatch.Elapsed:mm\\:ss}");

            List<ModelTableInfo> originalTables;
            List<TableSchema> dumpedTables;
            using (var model = new ModelQuery(workbook))
            {
                originalTables = ModelSchema.ReadTables(workbook, model, countRows: false);
                dumpedTables = DumpTables(model, originalTables, config, cacheFolder, reporter);
            }

            if (!config.PowerPivot.CacheSelfCheck || !hasParamsTable)
                return new CacheResult(dumpedTables, false, []);

            return new CacheResult(dumpedTables, true, RunSelfCheck(session, workbook, config, originalTables, dumpedTables, reporter));
        }
        finally
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
        }
    }

    private static void EnsureCachedQueriesHaveTables(SplitConfig config)
    {
        if (config.PowerPivot.CacheTables.Count == 0) return;

        var wantedTables = new HashSet<string>(config.PowerPivot.CacheTables, StringComparer.OrdinalIgnoreCase);
        var missing = config.PowerPivot.Queries
            .Where(query => query.Cache && !wantedTables.Contains(query.Name))
            .Select(query => query.Name)
            .ToList();

        if (missing.Count > 0)
            throw new InvalidOperationException(
                $"Zapytania z Cache = true, których tabel nie ma w CacheTables: {string.Join(", ", missing)}. " +
                "Bez ich plików CSV odświeżenie z cache się nie powiedzie.");
    }

    private static bool SetParametersForFullDump(Excel.Workbook workbook, SplitConfig config, string cacheFolder)
    {
        string cacheFolderParam = cacheFolder.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

        if (!WorkbookParams.TrySetParameter(workbook, config.PowerPivot.ParamsTable, config.PowerPivot.UseCacheParamName, "0")) return false;

        WorkbookParams.TrySetParameter(workbook, config.PowerPivot.ParamsTable, config.PowerPivot.SplitParamName, config.PowerPivot.AllValuesToken);
        WorkbookParams.TrySetParameter(workbook, config.PowerPivot.ParamsTable, config.PowerPivot.CachePathParamName, cacheFolderParam);
        return true;
    }

    private static List<TableSchema> DumpTables(
        ModelQuery model, List<ModelTableInfo> tables, SplitConfig config, string cacheFolder, IProgressReporter reporter)
    {
        var wantedTables = config.PowerPivot.CacheTables.Count > 0
            ? new HashSet<string>(config.PowerPivot.CacheTables, StringComparer.OrdinalIgnoreCase)
            : null;

        var dumped = new List<TableSchema>();
        foreach (ModelTableInfo table in tables)
        {
            if (wantedTables is not null && !wantedTables.Contains(table.Name)) continue;

            string csvPath = Path.Combine(cacheFolder, SanitizeFileName(table.Name) + ".csv");
            var timer = Stopwatch.StartNew();
            TableSchema schema = DumpTableToCsv(model, table, csvPath);
            WriteTypesFile(schema, Path.ChangeExtension(csvPath, ".types.csv"));
            dumped.Add(schema);

            reporter.Detail($"{table.Name,-24} {schema.RowCount,10:N0} w.  {timer.Elapsed:mm\\:ss}  -> {Path.GetFileName(csvPath)}");
            if (schema.SkippedCalculatedColumns.Length > 0)
                reporter.Detail($"  pominięte kolumny obliczeniowe (policzy je model): {string.Join(", ", schema.SkippedCalculatedColumns)}");
        }
        return dumped;
    }

    private static List<string> RunSelfCheck(
        ExcelSession session, Excel.Workbook workbook, SplitConfig config,
        List<ModelTableInfo> originalTables, List<TableSchema> dumpedTables, IProgressReporter reporter)
    {
        reporter.Step("Samokontrola: buduję model z cache i porównuję z oryginałem...");
        var stopwatch = Stopwatch.StartNew();
        var problems = new List<string>();

        WorkbookParams.SetParameter(workbook, config.PowerPivot.ParamsTable, config.PowerPivot.UseCacheParamName, "1");

        var refreshStarted = DateTime.Now;
        try
        {
            session.RefreshAllAndWait(workbook, config.PowerPivot.RefreshPivotCaches, reporter);
        }
        catch (Exception ex)
        {
            problems.Add($"odświeżenie z cache się nie powiodło: {ex.Message}");
            return problems;
        }

        var notRefreshed = ModelSchema.FindTablesNotRefreshed(
            refreshStarted, ModelSchema.ReadRefreshDates(workbook), dumpedTables.Select(schema => schema.TableName));
        foreach (StaleTable stale in notRefreshed)
            problems.Add($"tabela '{stale.Name}' nie odświeżyła się z cache - została z datą " +
                         $"{stale.RefreshedAt:yyyy-MM-dd HH\\:mm\\:ss}, a odświeżenie zaczęło się " +
                         $"{refreshStarted:yyyy-MM-dd HH\\:mm\\:ss}");

        using var model = new ModelQuery(workbook);
        problems.AddRange(ModelSchema.FindDifferences(originalTables, ModelSchema.ReadTables(workbook, model, countRows: false)));

        foreach (var schema in dumpedTables)
        {
            long rowCountFromCache = Com.Or(() => model.CountTableRows(schema.TableName), 0L);
            if (rowCountFromCache != schema.RowCount)
                problems.Add($"'{schema.TableName}': {rowCountFromCache:N0} wierszy z cache zamiast {schema.RowCount:N0}");
        }

        reporter.Detail($"samokontrola: {stopwatch.Elapsed:mm\\:ss}");
        return problems;
    }

    private static TableSchema DumpTableToCsv(ModelQuery model, ModelTableInfo table, string csvPath)
    {
        var dataColumns = table.Columns.Where(column => !column.IsCalculated).ToList();
        string[] skippedColumns = table.Columns.Where(column => column.IsCalculated).Select(column => column.Name).ToArray();
        if (dataColumns.Count == 0)
            throw new InvalidOperationException($"Tabela '{table.Name}' nie ma kolumn z danymi - nie ma czego zrzucić.");

        string dax = "EVALUATE SELECTCOLUMNS(" + ModelQuery.TableRef(table.Name) + ", " +
                     string.Join(", ", dataColumns.Select(column =>
                         ModelQuery.TextLiteral(column.Name) + ", " + ModelQuery.ColumnRef(table.Name, column.Name))) +
                     ")";

        var columnKinds = new ColumnKind[dataColumns.Count];
        var allValuesAtMidnight = Enumerable.Repeat(true, dataColumns.Count).ToArray();
        long rowCount = 0;

        using (var writer = new StreamWriter(csvPath, false, new UTF8Encoding(true)))
        {
            model.StreamRows(
                dax,
                header =>
                {
                    if (header.Length != dataColumns.Count)
                        throw new InvalidOperationException($"'{table.Name}': DAX zwrócił {header.Length} kolumn zamiast {dataColumns.Count}.");
                    writer.WriteLine(string.Join(";", dataColumns.Select(column => EscapeCsv(column.Name))));
                },
                row =>
                {
                    rowCount++;
                    var cells = new string[row.Length];
                    for (int i = 0; i < row.Length; i++)
                    {
                        columnKinds[i] = MergeKinds(columnKinds[i], ClassifyValue(row[i]));
                        if (row[i] is DateTime value && value.TimeOfDay != TimeSpan.Zero) allValuesAtMidnight[i] = false;
                        cells[i] = FormatCsvValue(row[i]);
                    }
                    writer.WriteLine(string.Join(";", cells));
                });
        }

        var columnTypes = dataColumns
            .Select((column, i) => CacheTypeName(column.DataType, columnKinds[i], allValuesAtMidnight[i]))
            .ToArray();
        return new TableSchema(
            table.Name, csvPath, dataColumns.Select(column => column.Name).ToArray(), columnTypes, rowCount, skippedColumns);
    }

    private static string CacheTypeName(int adoType, ColumnKind inferredKind, bool allValuesAtMidnight) => adoType switch
    {
        130 or 129 or 200 or 201 or 202 or 203 => "text",
        7 or 133 or 134 or 135 => allValuesAtMidnight ? "date" : "datetime",
        2 or 3 or 16 or 17 or 18 or 19 or 20 or 21 => "int",
        4 or 5 or 14 or 131 => "number",
        6 => "currency",
        11 => "logical",
        _ => inferredKind switch
        {
            ColumnKind.Int => "int",
            ColumnKind.Number => "number",
            ColumnKind.DateTime => allValuesAtMidnight ? "date" : "datetime",
            ColumnKind.Logical => "logical",
            _ => "text",
        },
    };

    private static void WriteTypesFile(TableSchema schema, string path)
    {
        var content = new StringBuilder();
        content.AppendLine("Column;Type");
        for (int i = 0; i < schema.ColumnNames.Length; i++)
            content.AppendLine($"{EscapeCsv(schema.ColumnNames[i])};{schema.ColumnTypes[i]}");
        File.WriteAllText(path, content.ToString(), new UTF8Encoding(true));
    }

    private static string FormatCsvValue(object? value) => value switch
    {
        null => "",
        DateTime date => date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        bool flag => flag ? "true" : "false",
        double number => number.ToString(CultureInfo.InvariantCulture),
        float number => number.ToString(CultureInfo.InvariantCulture),
        decimal number => number.ToString(CultureInfo.InvariantCulture),
        byte or sbyte or short or ushort or int or uint or long or ulong
            => Convert.ToString(value, CultureInfo.InvariantCulture)!,
        _ => EscapeCsv(Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""),
    };

    private static string EscapeCsv(string text)
    {
        if (text.IndexOfAny([';', '"', '\n', '\r']) < 0) return text;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    private static ColumnKind ClassifyValue(object? value) => value switch
    {
        null => ColumnKind.Unknown,
        DateTime => ColumnKind.DateTime,
        bool => ColumnKind.Logical,
        double or float or decimal => ColumnKind.Number,
        byte or sbyte or short or ushort or int or uint or long or ulong => ColumnKind.Int,
        _ => ColumnKind.Text,
    };

    private static ColumnKind MergeKinds(ColumnKind first, ColumnKind second)
    {
        if (first == ColumnKind.Unknown) return second;
        if (second == ColumnKind.Unknown) return first;
        if (first == second) return first;
        if ((first is ColumnKind.Int or ColumnKind.Number) && (second is ColumnKind.Int or ColumnKind.Number)) return ColumnKind.Number;
        return ColumnKind.Text;
    }

    private static string SanitizeFileName(string name) =>
        string.Join("_", name.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));
}
