namespace ExcelFileSplitter.Core;

public static class Preflight
{
    public static List<string> FindProblems(WorkbookInfo info, SplitConfig config, IReadOnlyList<Recipient> recipients)
    {
        var problems = new List<string>();

        if (info.OutdatedQueries.Count > 0)
            problems.Add("skoroszyt nie jest przygotowany bieżącą wersją aplikacji (zapytania: " +
                         string.Join(", ", info.OutdatedQueries) + ") - uruchom `prepare`, a potem `cache`");

        problems.AddRange(info.DataProblems);

        foreach (SheetSet set in DistinctSheetSets(config, recipients))
            foreach (string problem in SheetDependencies.FindSheetProblems(info, config, set.Sheets))
                problems.Add(set.Label is null ? problem : $"{set.Label} - {problem}");

        return problems;
    }

    private sealed record SheetSet(IReadOnlyList<string> Sheets, string? Label);

    private static List<SheetSet> DistinctSheetSets(SplitConfig config, IReadOnlyList<Recipient> recipients)
    {
        if (recipients.Count == 0) return [new SheetSet(config.SheetsToKeep, null)];

        var bySheets = recipients
            .GroupBy(
                recipient => string.Join('\u0001', recipient.SheetsToKeep.OrderBy(sheet => sheet, StringComparer.OrdinalIgnoreCase)),
                StringComparer.OrdinalIgnoreCase)
            .ToList();

        return bySheets
            .Select(group => new SheetSet(group.First().SheetsToKeep, bySheets.Count == 1 ? null : DescribeFiles(group)))
            .ToList();
    }

    private static string DescribeFiles(IEnumerable<Recipient> recipients)
    {
        List<string> names = recipients.Select(recipient => $"'{recipient.Name}'").ToList();
        return names.Count <= 3
            ? "pliki " + string.Join(", ", names)
            : $"pliki {string.Join(", ", names.Take(3))} i {names.Count - 3} inn.";
    }
}
