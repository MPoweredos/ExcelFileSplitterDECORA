using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public static class PowerQuery
{
    public static List<string> ListQueryNames(Excel.Workbook workbook)
    {
        var names = new List<string>();
        Com.Try(() =>
        {
            foreach (Excel.WorkbookQuery query in workbook.Queries) names.Add(query.Name);
        });
        return names;
    }

    public static bool QueryExists(Excel.Workbook workbook, string name) =>
        ListQueryNames(workbook).Any(existing => string.Equals(existing, name, StringComparison.OrdinalIgnoreCase));

    public static string GetFormula(Excel.Workbook workbook, string name) =>
        FindQuery(workbook, name)?.Formula
        ?? throw new InvalidOperationException($"W skoroszycie nie ma zapytania '{name}'.");

    public static void SetFormula(Excel.Workbook workbook, string name, string formula)
    {
        Excel.WorkbookQuery query = FindQuery(workbook, name)
            ?? throw new InvalidOperationException($"W skoroszycie nie ma zapytania '{name}'.");
        query.Formula = formula;
    }

    public static void AddConnectionOnlyQuery(Excel.Workbook workbook, string name, string formula, string description) =>
        workbook.Queries.Add(name, formula, description);

    public static int ReplaceAllFormulasWithError(Excel.Workbook workbook, string message, IProgressReporter reporter)
    {
        int replaced = 0;
        var failed = new List<string>();

        foreach (Excel.WorkbookQuery query in AllQueries(workbook))
        {
            string name = Com.Or(() => query.Name, "(bez nazwy)");
            bool done = Com.Try(() =>
                query.Formula = $"let Source = error \"{message.Replace("\"", "'")}\" in Source");

            if (done) replaced++;
            else failed.Add(name);
        }

        if (failed.Count > 0)
            reporter.Warning($"nie udało się zaślepić zapytań: {string.Join(", ", failed)} - zostają w pliku " +
                             "razem ze ścieżkami do źródeł");

        return replaced;
    }

    public static int DeleteAllQueries(Excel.Workbook workbook, IProgressReporter reporter)
    {
        int deleted = 0;
        var failed = new List<string>();

        foreach (Excel.WorkbookQuery query in AllQueries(workbook))
        {
            string name = Com.Or(() => query.Name, "(bez nazwy)");
            if (Com.Try(() => query.Delete())) deleted++;
            else failed.Add(name);
        }

        if (failed.Count > 0)
            reporter.Warning($"nie udało się usunąć zapytań: {string.Join(", ", failed)}");

        return deleted;
    }

    private static Excel.WorkbookQuery? FindQuery(Excel.Workbook workbook, string name)
    {
        foreach (Excel.WorkbookQuery query in workbook.Queries)
            if (string.Equals(query.Name, name, StringComparison.OrdinalIgnoreCase))
                return query;
        return null;
    }

    private static List<Excel.WorkbookQuery> AllQueries(Excel.Workbook workbook)
    {
        var queries = new List<Excel.WorkbookQuery>();
        Com.Try(() =>
        {
            foreach (Excel.WorkbookQuery query in workbook.Queries) queries.Add(query);
        });
        return queries;
    }
}
