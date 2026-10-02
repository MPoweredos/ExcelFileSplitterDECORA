using System.Text.RegularExpressions;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public enum DependencyKind { Chart, PivotChart, Formula, Pivot }

public sealed record SheetDependency(string Sheet, string Requires, DependencyKind Kind, string Detail);

public static class SheetDependencies
{
    private static readonly Regex SheetRefPattern =
        new(@"(?:'(?<q>(?:[^']|'')+)'|(?<p>[\p{L}\p{N}_.]+))!", RegexOptions.Compiled);

    public static Dictionary<string, string> FindSheetsWithModelTables(Excel.Workbook workbook)
    {
        var modelTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Com.Try(() =>
        {
            foreach (Excel.ModelTable modelTable in workbook.Model.ModelTables) modelTableNames.Add(modelTable.Name);
        });

        var sheetsWithModelTables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            foreach (Excel.ListObject table in sheet.ListObjects)
                if (modelTableNames.Contains(table.Name) && !sheetsWithModelTables.ContainsKey(sheet.Name))
                    sheetsWithModelTables[sheet.Name] = table.Name;
        return sheetsWithModelTables;
    }

    public static List<SheetDependency> FindDependencies(Excel.Workbook workbook)
    {
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Excel.Worksheet sheet in workbook.Worksheets) sheetNames.Add(sheet.Name);

        Dictionary<string, string> sheetsByTable = FindSheetsByTable(workbook);

        var dependencies = new List<SheetDependency>();
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            CollectChartDependencies(sheet, sheetNames, dependencies);
            CollectPivotDependencies(sheet, sheetNames, sheetsByTable, dependencies);
            CollectFormulaDependencies(sheet, sheetNames.ToList(), dependencies);
        }
        return dependencies
            .DistinctBy(dependency => (dependency.Sheet, dependency.Requires, dependency.Kind, dependency.Detail))
            .ToList();
    }

    public static List<string> FindSheetProblems(WorkbookInfo info, SplitConfig config, IReadOnlyList<string> sheetsToKeep)
    {
        var keptSheets = new HashSet<string>(sheetsToKeep, StringComparer.OrdinalIgnoreCase);
        var problems = new List<string>();

        var missingSheets = sheetsToKeep.Where(kept => !info.Sheets.Contains(kept, StringComparer.OrdinalIgnoreCase)).ToList();
        if (missingSheets.Count > 0)
            problems.Add("w skoroszycie nie ma arkuszy: " + string.Join(", ", missingSheets.Select(Quote)));

        if (keptSheets.Contains(config.PowerPivot.ConfigSheetName))
            problems.Add($"{Quote(config.PowerPivot.ConfigSheetName)} to arkusz parametrów podziału (m.in. ścieżka do cache) - " +
                         "nie może trafić do odbiorcy");

        foreach (var (sheet, table) in info.SheetsWithModelTables)
            if (!keptSheets.Contains(sheet))
                problems.Add($"{Quote(sheet)} trzyma tabelę {Quote(table)} załadowaną do modelu danych - bez tego arkusza " +
                             "tabela zniknie z modelu, a miary przestaną działać");

        foreach (var dependency in info.Dependencies)
        {
            if (!keptSheets.Contains(dependency.Sheet) || keptSheets.Contains(dependency.Requires)) continue;

            if (dependency.Kind == DependencyKind.Pivot
                && info.DataSheetName is not null
                && SameName(dependency.Requires, info.DataSheetName))
                continue;

            problems.Add(DescribeProblem(dependency));
        }

        if (missingSheets.Count == 0 && keptSheets.Count > 0 && !info.VisibleSheets.Any(keptSheets.Contains))
            problems.Add("wszystkie wybrane arkusze są w źródle ukryte - Excel wymaga co najmniej jednego widocznego");

        return problems;
    }

    private static void CollectChartDependencies(Excel.Worksheet sheet, HashSet<string> sheetNames, List<SheetDependency> dependencies)
    {
        var chartObjects = Com.Or<Excel.ChartObjects?>(() => (Excel.ChartObjects)sheet.ChartObjects(), null);
        if (chartObjects is null) return;

        for (int i = 1; i <= chartObjects.Count; i++)
        {
            int index = i;
            var chartObject = Com.Or<Excel.ChartObject?>(() => (Excel.ChartObject)chartObjects.Item(index), null);
            if (chartObject is null) continue;

            string chartName = chartObject.Name;
            Excel.Chart chart = chartObject.Chart;

            string? pivotSheetName = null;
            try
            {
                Excel.PivotTable pivotTable = chart.PivotLayout.PivotTable;
                pivotSheetName = ((Excel.Worksheet)pivotTable.Parent).Name;
                if (!SameName(pivotSheetName, sheet.Name))
                    dependencies.Add(new SheetDependency(sheet.Name, pivotSheetName, DependencyKind.PivotChart,
                        $"wykres '{chartName}' (tabela przestawna '{pivotTable.Name}')"));
            }
            catch (Exception exception) when (Com.SaysNo(exception))
            {
            }

            var seriesCollection = Com.Or<Excel.SeriesCollection?>(() => (Excel.SeriesCollection)chart.SeriesCollection(), null);
            if (seriesCollection is null) continue;

            for (int seriesIndex = 1; seriesIndex <= seriesCollection.Count; seriesIndex++)
            {
                int series = seriesIndex;
                string? seriesFormula = Com.Or<string?>(
                    () => (string)((Excel.Series)seriesCollection.Item(series)).Formula, null);
                if (seriesFormula is null) continue;

                foreach (string required in ReferencedSheets(seriesFormula, sheetNames))
                    if (!SameName(required, sheet.Name) && (pivotSheetName is null || !SameName(required, pivotSheetName)))
                        dependencies.Add(new SheetDependency(sheet.Name, required, DependencyKind.Chart, $"wykres '{chartName}'"));
            }
        }
    }

    private static Dictionary<string, string> FindSheetsByTable(Excel.Workbook workbook)
    {
        var sheetsByTable = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
            foreach (Excel.ListObject table in sheet.ListObjects)
                sheetsByTable[table.Name] = sheet.Name;
        return sheetsByTable;
    }

    private static void CollectPivotDependencies(
        Excel.Worksheet sheet, HashSet<string> sheetNames, Dictionary<string, string> sheetsByTable,
        List<SheetDependency> dependencies)
    {
        var pivotTables = Com.Or<Excel.PivotTables?>(() => (Excel.PivotTables)sheet.PivotTables(), null);
        if (pivotTables is null) return;

        for (int i = 1; i <= pivotTables.Count; i++)
        {
            Excel.PivotTable pivotTable;
            string sourceData;
            try
            {
                pivotTable = (Excel.PivotTable)pivotTables.Item(i);
                sourceData = Convert.ToString(pivotTable.PivotCache().SourceData) ?? "";
            }
            catch (Exception exception) when (Com.SaysNo(exception))
            {
                continue;
            }

            string detail = $"tabela przestawna '{pivotTable.Name}'";

            bool found = false;
            foreach (string required in ReferencedSheets(sourceData, sheetNames))
            {
                found = true;
                if (!SameName(required, sheet.Name))
                    dependencies.Add(new SheetDependency(sheet.Name, required, DependencyKind.Pivot, detail));
            }

            if (!found && sheetsByTable.TryGetValue(sourceData.Trim(), out string? tableSheet) && !SameName(tableSheet, sheet.Name))
                dependencies.Add(new SheetDependency(sheet.Name, tableSheet, DependencyKind.Pivot, detail));
        }
    }

    private static void CollectFormulaDependencies(
        Excel.Worksheet sheet, List<string> sheetNames, List<SheetDependency> dependencies)
    {
        var formulaCells = Com.Or<Excel.Range?>(
            () => sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas), null);
        if (formulaCells is null) return;

        List<string> looking = sheetNames.Where(name => !SameName(name, sheet.Name)).ToList();

        foreach (Excel.Range area in formulaCells.Areas)
        {
            if (looking.Count == 0) return;

            for (int i = looking.Count - 1; i >= 0; i--)
            {
                Excel.Range? hit = FindReferenceTo(area, looking[i]);
                if (hit is null) continue;

                string formula = Com.Or(() => (string)hit.Formula, "");
                dependencies.Add(new SheetDependency(
                    sheet.Name, looking[i], DependencyKind.Formula, "formula " + Truncate(formula, 60)));
                looking.RemoveAt(i);
            }
        }
    }

    private static Excel.Range? FindReferenceTo(Excel.Range area, string sheetName)
    {
        string[] needles =
        [
            ForFind(sheetName) + "!",
            "'" + ForFind(sheetName.Replace("'", "''")) + "'!",
        ];

        foreach (string needle in needles)
        {
            var hit = Com.Or<Excel.Range?>(() => area.Find(
                What: needle,
                LookIn: Excel.XlFindLookIn.xlFormulas,
                LookAt: Excel.XlLookAt.xlPart,
                MatchCase: false), null);

            if (hit is not null) return hit;
        }
        return null;
    }

    private static string ForFind(string text) =>
        text.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");

    private static IEnumerable<string> ReferencedSheets(string formula, HashSet<string> sheetNames)
    {
        foreach (Match match in SheetRefPattern.Matches(formula))
        {
            string name = match.Groups["q"].Success ? match.Groups["q"].Value.Replace("''", "'") : match.Groups["p"].Value;
            if (sheetNames.TryGetValue(name, out string? actualName)) yield return actualName;
        }
    }

    private static string DescribeProblem(SheetDependency dependency) => dependency.Kind switch
    {
        DependencyKind.Chart =>
            $"{Quote(dependency.Sheet)}: {dependency.Detail} pobiera dane z {Quote(dependency.Requires)} - po jego skasowaniu wykres zamrozi się " +
            "na danych sprzed podziału, czyli z danymi innych odbiorców",
        DependencyKind.PivotChart =>
            $"{Quote(dependency.Sheet)}: {dependency.Detail} stoi na {Quote(dependency.Requires)} - po jego skasowaniu wykres zamrozi się " +
            "na danych sprzed podziału, czyli z danymi innych odbiorców",
        DependencyKind.Pivot =>
            $"{Quote(dependency.Sheet)}: {dependency.Detail} czyta dane z {Quote(dependency.Requires)} - po jego skasowaniu zostanie " +
            "z własnym cache'em, czyli z kompletem wierszy sprzed podziału",
        _ =>
            $"{Quote(dependency.Sheet)}: {dependency.Detail} odwołuje się do {Quote(dependency.Requires)} - po jego skasowaniu formuły dostaną #REF!",
    };

    private static bool SameName(string first, string second) => string.Equals(first, second, StringComparison.OrdinalIgnoreCase);

    private static string Truncate(string text, int maxLength) => text.Length <= maxLength ? text : text[..maxLength] + "...";

    private static string Quote(string text) => "'" + text + "'";
}
