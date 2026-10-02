using System.Text;
using System.Text.RegularExpressions;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public static class WorkbookPreparer
{
    public const string Version = "v3";

    private static readonly string WrapperHeader =
        $"// [ExcelFileSplitter {Version}] zapytanie przygotowane automatycznie - nie edytuj tego nagłówka";

    public static bool IsPreparedByCurrentVersion(string formula) =>
        formula.Contains($"[ExcelFileSplitter {Version}]", StringComparison.Ordinal);

    public static string PrepareWorkbook(ExcelSession session, SplitConfig config, IProgressReporter reporter)
    {
        if (config.PowerPivot.Queries.Count == 0)
            throw new InvalidOperationException("Sekcja 'Queries' jest pusta - nie ma czego przygotować.");

        string preparedPath = config.PreparedWorkbookPath();
        Directory.CreateDirectory(Path.GetDirectoryName(preparedPath)!);
        File.Copy(Path.GetFullPath(config.SourceWorkbook), preparedPath, overwrite: true);
        File.SetAttributes(preparedPath, FileAttributes.Normal);

        Excel.Workbook? workbook = null;
        try
        {
            workbook = session.OpenWorkbook(preparedPath);

            CreateParamsSheet(workbook, config);
            reporter.Step($"Arkusz '{config.PowerPivot.ConfigSheetName}' z tabela '{config.PowerPivot.ParamsTable}'");

            AddHelperQueries(workbook, config, reporter);

            foreach (QuerySpec querySpec in config.PowerPivot.Queries)
                InjectFilterIntoQuery(workbook, config, querySpec, reporter);

            workbook.Save();
            return preparedPath;
        }
        finally
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
        }
    }

    private static void CreateParamsSheet(Excel.Workbook workbook, SplitConfig config)
    {
        foreach (Excel.Worksheet existingSheet in workbook.Worksheets)
        {
            if (string.Equals(existingSheet.Name, config.PowerPivot.ConfigSheetName, StringComparison.OrdinalIgnoreCase))
            {
                existingSheet.Delete();
                break;
            }
        }

        var sheet = (Excel.Worksheet)workbook.Worksheets.Add();
        sheet.Name = config.PowerPivot.ConfigSheetName;

        string cacheFolder = string.IsNullOrWhiteSpace(config.PowerPivot.CacheFolder)
            ? ""
            : Path.GetFullPath(config.PowerPivot.CacheFolder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

        ((Excel.Range)sheet.Cells[1, 1]).Value2 = WorkbookParams.ParameterColumnHeader;
        ((Excel.Range)sheet.Cells[1, 2]).Value2 = WorkbookParams.ValueColumnHeader;

        var parameterRows = new (string Key, string Value)[]
        {
            (config.PowerPivot.SplitParamName, config.PowerPivot.AllValuesToken),
            (config.PowerPivot.UseCacheParamName, "0"),
            (config.PowerPivot.CachePathParamName, cacheFolder),
        };

        for (int i = 0; i < parameterRows.Length; i++)
        {
            ((Excel.Range)sheet.Cells[i + 2, 1]).Value2 = parameterRows[i].Key;
            var valueCell = (Excel.Range)sheet.Cells[i + 2, 2];
            valueCell.NumberFormat = "@";
            valueCell.Value2 = parameterRows[i].Value;
        }

        Excel.Range tableRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[parameterRows.Length + 1, 2]];
        Excel.ListObject paramsTable = sheet.ListObjects.Add(
            Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes);
        paramsTable.Name = config.PowerPivot.ParamsTable;

        ((Excel.Range)sheet.Columns[1]).ColumnWidth = 18;
        ((Excel.Range)sheet.Columns[2]).ColumnWidth = 60;
    }

    private static void AddHelperQueries(Excel.Workbook workbook, SplitConfig config, IProgressReporter reporter)
    {
        AddOrReplaceQuery(workbook, "fnParam", FnParamFormula(config),
            "ExcelFileSplitter: odczyt parametru z tabeli " + config.PowerPivot.ParamsTable, reporter);

        if (config.PowerPivot.Queries.Any(query => query.Cache))
            AddOrReplaceQuery(workbook, "fnCache", FnCacheFormula(config),
                "ExcelFileSplitter: odczyt tabeli z cache CSV", reporter);
    }

    private static void AddOrReplaceQuery(
        Excel.Workbook workbook, string name, string formula, string description, IProgressReporter reporter)
    {
        if (PowerQuery.QueryExists(workbook, name))
        {
            PowerQuery.SetFormula(workbook, name, formula);
            reporter.Step($"Zapytanie '{name}' zaktualizowane");
        }
        else
        {
            PowerQuery.AddConnectionOnlyQuery(workbook, name, formula, description);
            reporter.Step($"Zapytanie '{name}' dodane");
        }
    }

    private static string FnParamFormula(SplitConfig config) => $$"""
        let
            fnParam = (name as text) as text =>
                let
                    ParamsTable = Excel.CurrentWorkbook(){[Name="{{config.PowerPivot.ParamsTable}}"]}[Content],
                    MatchingRow = Table.SelectRows(ParamsTable, each [{{WorkbookParams.ParameterColumnHeader}}] = name),
                    Result      = if Table.IsEmpty(MatchingRow) then "" else Text.From(MatchingRow{0}[{{WorkbookParams.ValueColumnHeader}}])
                in
                    Result
        in
            fnParam
        """;

    private static string FnCacheFormula(SplitConfig config) => $$"""
        let
            fnCache = (tableName as text) as table =>
                let
                    CacheFolder = fnParam("{{config.PowerPivot.CachePathParamName}}"),
                    CsvOptions  = [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv],
                    RawTable    = Table.PromoteHeaders(
                                      Csv.Document(File.Contents(CacheFolder & tableName & ".csv"), CsvOptions),
                                      [PromoteAllScalars=true]),
                    SchemaRows  = Table.ToRecords(Table.PromoteHeaders(
                                      Csv.Document(File.Contents(CacheFolder & tableName & ".types.csv"), CsvOptions),
                                      [PromoteAllScalars=true])),
                    ToType      = (typeName as text) as type =>
                                      if typeName = "int"           then Int64.Type
                                      else if typeName = "number"   then type number
                                      else if typeName = "currency" then Currency.Type
                                      else if typeName = "date"     then type datetime
                                      else if typeName = "datetime" then type datetime
                                      else if typeName = "logical"  then type logical
                                      else type text,
                    SelectedColumns = Table.SelectColumns(RawTable, List.Transform(SchemaRows, each [Column])),
                    TypedTable      = Table.TransformColumnTypes(
                                          SelectedColumns,
                                          List.Transform(SchemaRows, each {[Column], ToType([Type])}), "en-US"),
                    // Daty są w CSV z godziną, więc do typu date schodzimy przez datetime.
                    Result          = Table.TransformColumnTypes(TypedTable,
                                          List.Transform(List.Select(SchemaRows, each [Type] = "date"), each {[Column], type date}))
                in
                    Result
        in
            fnCache
        """;

    private static void InjectFilterIntoQuery(
        Excel.Workbook workbook, SplitConfig config, QuerySpec querySpec, IProgressReporter reporter)
    {
        string originalFormula = PowerQuery.GetFormula(workbook, querySpec.Name);

        if (originalFormula.Contains("[ExcelFileSplitter", StringComparison.Ordinal))
            throw new InvalidOperationException(
                $"Zapytanie '{querySpec.Name}' w SourceWorkbook jest już przygotowane. SourceWorkbook musi wskazywać " +
                "oryginalny raport, nie plik .prepared.");

        PowerQuery.SetFormula(workbook, querySpec.Name, BuildWrappedFormula(config, querySpec, originalFormula));

        string filterDescription = querySpec.KeyFrom is not null
            ? $"filtr przez klucz {querySpec.KeyFrom.Query}[{querySpec.KeyFrom.Column}]"
            : querySpec.FilterColumn is { Length: > 0 } ? $"filtr po [{querySpec.FilterColumn}]" : "bez filtru";
        reporter.Step($"'{querySpec.Name}': {filterDescription}{(querySpec.Cache ? " + cache" : "")}");
    }

    internal static string BuildWrappedFormula(SplitConfig config, QuerySpec querySpec, string originalFormula)
    {
        var formula = new StringBuilder();
        formula.AppendLine(WrapperHeader);
        formula.AppendLine("let");
        formula.AppendLine($"    __splitValue = fnParam({MText(config.PowerPivot.SplitParamName)}),");

        if (querySpec.Cache)
        {
            formula.AppendLine($"    __source = if fnParam({MText(config.PowerPivot.UseCacheParamName)}) = \"1\" then fnCache({MText(querySpec.Name)}) else (");
            formula.AppendLine(IndentBlock(originalFormula));
            formula.AppendLine("    ),");
        }
        else
        {
            formula.AppendLine("    __source = (");
            formula.AppendLine(IndentBlock(originalFormula));
            formula.AppendLine("    ),");
        }

        formula.AppendLine("    __result =");
        formula.AppendLine(BuildFilterExpression(config, querySpec));
        formula.AppendLine("in");
        formula.Append("    __result");
        return formula.ToString();
    }

    private static string BuildFilterExpression(SplitConfig config, QuerySpec querySpec)
    {
        if (string.IsNullOrWhiteSpace(querySpec.FilterColumn))
            return "        __source";

        string allValuesToken = MText(config.PowerPivot.AllValuesToken);
        string filterColumn = MText(querySpec.FilterColumn!);

        if (querySpec.KeyFrom is null)
        {
            return
                "        if __splitValue = " + allValuesToken + " then __source\r\n" +
                "        else\r\n" +
                "            let\r\n" +
                "                __values = List.Buffer(Text.Split(Text.Replace(__splitValue, \"#(cr)\", \"\"), \"#(lf)\"))\r\n" +
                "            in\r\n" +
                "                Table.SelectRows(__source, each List.Contains(__values, Record.Field(_, " + filterColumn + ")))";
        }

        string dimensionQuery = MIdentifier(querySpec.KeyFrom.Query);
        string keyColumn = MText(querySpec.KeyFrom.Column);

        return
            "        if __splitValue = " + allValuesToken + " then __source\r\n" +
            "        else\r\n" +
            "            let\r\n" +
            "                __keys = Table.RenameColumns(\r\n" +
            "                    Table.Distinct(Table.SelectColumns(" + dimensionQuery + ", {" + keyColumn + "})),\r\n" +
            "                    {{" + keyColumn + ", \"__efs_key\"}}),\r\n" +
            "                __joined = Table.Join(__source, {" + filterColumn + "}, __keys, {\"__efs_key\"}, JoinKind.Inner)\r\n" +
            "            in\r\n" +
            "                Table.RemoveColumns(__joined, {\"__efs_key\"})";
    }

    private static string IndentBlock(string text) =>
        string.Join("\r\n", text.Replace("\r\n", "\n").Split('\n').Select(line => "        " + line));

    private static string MText(string text) => "\"" + text.Replace("\"", "\"\"") + "\"";

    private static string MIdentifier(string name) =>
        Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$") ? name : "#\"" + name.Replace("\"", "\"\"") + "\"";
}
