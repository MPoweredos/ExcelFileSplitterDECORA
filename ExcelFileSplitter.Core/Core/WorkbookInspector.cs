using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed record WorkbookInfo(
    List<string> Sheets,
    List<string> VisibleSheets,
    List<ModelTableInfo> ModelTables,
    List<string> QueryNames,
    List<SplitKey> SplitKeys,
    Dictionary<string, string> SheetsWithModelTables,
    List<SheetDependency> Dependencies,
    List<string> OutdatedQueries,
    List<string> DataProblems,
    string? DataSheetName);

public static class WorkbookInspector
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
        bool readSplitValues = true)
    {
        List<SheetInfo> sheets = Interop.Sheets.List(workbook);
        List<string> queryNames = PowerQuery.ListQueryNames(workbook);
        Dictionary<string, string> sheetsWithModelTables = SheetDependencies.FindSheetsWithModelTables(workbook);
        List<SheetDependency> dependencies = SheetDependencies.FindDependencies(workbook);
        List<string> outdatedQueries = FindOutdatedQueries(workbook, config, queryNames);

        List<ModelTableInfo> modelTables;
        var splitKeys = new List<SplitKey>();

        using (var model = new ModelQuery(workbook))
        {
            modelTables = ModelSchema.ReadTables(workbook, model, countRows: true);

            if (readSplitValues && !string.IsNullOrWhiteSpace(config.PowerPivot.SplitTable) && !string.IsNullOrWhiteSpace(config.SplitColumn))
            {
                splitKeys = ReadSplitColumnValues(model, config, reporter).Select(value => SplitKey.Of(value)).ToList();
            }
        }

        return new WorkbookInfo(
            sheets.Select(sheet => sheet.Name).ToList(),
            sheets.Where(sheet => sheet.IsVisible).Select(sheet => sheet.Name).ToList(),
            modelTables,
            queryNames,
            splitKeys,
            sheetsWithModelTables,
            dependencies,
            outdatedQueries,
            [],
            null);
    }

    private static List<string> FindOutdatedQueries(Excel.Workbook workbook, SplitConfig config, List<string> queryNames)
    {
        var outdated = new List<string>();
        foreach (QuerySpec querySpec in config.PowerPivot.Queries)
        {
            if (!queryNames.Contains(querySpec.Name, StringComparer.OrdinalIgnoreCase))
            {
                outdated.Add($"{querySpec.Name} (brak zapytania)");
                continue;
            }
            try
            {
                if (!WorkbookPreparer.IsPreparedByCurrentVersion(PowerQuery.GetFormula(workbook, querySpec.Name)))
                    outdated.Add(querySpec.Name);
            }
            catch { outdated.Add(querySpec.Name); }
        }
        return outdated;
    }

    private static List<string> ReadSplitColumnValues(ModelQuery model, SplitConfig config, IProgressReporter reporter)
    {
        try
        {
            return model.ReadDistinctValues(config.PowerPivot.SplitTable, config.SplitColumn);
        }
        catch (Exception ex)
        {
            reporter.Warning($"nie udało się odczytać wartości {config.PowerPivot.SplitTable}[{config.SplitColumn}]: {ex.Message}");
            return [];
        }
    }
}
