namespace ExcelFileSplitter.Core;

public sealed record Recipient(string Name, IReadOnlyList<SplitKey> Keys, IReadOnlyList<string> SheetsToKeep)
{
    public bool IsGroup => Keys.Count != 1 || !string.Equals(Keys[0].Display, Name, StringComparison.Ordinal);

    public IEnumerable<string> KeyLabels => Keys.Select(key => key.Display);
}

public sealed record RecipientPlan(List<Recipient> Files, List<string> Problems, List<string> Warnings);

public static class Recipients
{
    public static RecipientPlan BuildPlan(SplitConfig config, IReadOnlyList<SplitKey> keysInData)
    {
        var plan = new RecipientPlan([], [], []);
        Dictionary<SplitKey, SplitKey> spelling = SpellingFromData(keysInData);

        if (config.Groups.Count > 0) AddGroupFiles(config, spelling, plan);
        else AddKeyFiles(config, keysInData, spelling, plan);

        ReportValuesWithNewLine(plan);
        ReportCollidingFileNames(config, plan);
        return plan;
    }

    private static Dictionary<SplitKey, SplitKey> SpellingFromData(IReadOnlyList<SplitKey> keysInData)
    {
        var spelling = new Dictionary<SplitKey, SplitKey>();
        foreach (SplitKey key in keysInData) spelling.TryAdd(key, key);
        return spelling;
    }

    private static void AddGroupFiles(SplitConfig config, Dictionary<SplitKey, SplitKey> spelling, RecipientPlan plan)
    {
        foreach (RecipientGroup group in config.Groups)
        {
            if (string.IsNullOrWhiteSpace(group.Name))
            {
                plan.Problems.Add("grupa bez nazwy (Name) - nazwa trafia do nazwy pliku");
                continue;
            }

            List<SplitKey> requested = group.RequestedKeys();
            if (requested.Count == 0)
            {
                plan.Problems.Add($"grupa '{group.Name}' nie ma żadnych wartości (Values albo Keys)");
                continue;
            }

            plan.Files.Add(Resolve(
                group.Name.Trim(), requested, MergeSheets(config.SheetsToKeep, group.SheetsToKeep), spelling, plan));
        }

        ReportKeysInManyFiles(plan);
    }

    public static List<string> MergeSheets(IReadOnlyList<string> baseSheets, IReadOnlyList<string> extraSheets)
    {
        var merged = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string sheet in baseSheets.Concat(extraSheets))
        {
            string trimmed = sheet.Trim();
            if (trimmed.Length > 0 && seen.Add(trimmed)) merged.Add(trimmed);
        }
        return merged;
    }

    private static void AddKeyFiles(
        SplitConfig config, IReadOnlyList<SplitKey> keysInData, Dictionary<SplitKey, SplitKey> spelling, RecipientPlan plan)
    {
        List<SplitKey> requested = config.RequestedKeys();
        IReadOnlyList<SplitKey> keys = requested.Count > 0 ? requested : keysInData;

        List<string> sheets = MergeSheets(config.SheetsToKeep, []);
        var alreadyAdded = new HashSet<SplitKey>();

        foreach (SplitKey key in keys)
        {
            if (!alreadyAdded.Add(key)) continue;

            Recipient recipient = Resolve(key.Display, [key], sheets, spelling, plan);
            plan.Files.Add(recipient with { Name = recipient.Keys[0].Display });
        }
    }

    private static Recipient Resolve(
        string name, IEnumerable<SplitKey> requestedKeys, IReadOnlyList<string> sheetsToKeep,
        Dictionary<SplitKey, SplitKey> spelling, RecipientPlan plan)
    {
        var resolved = new List<SplitKey>();
        var notInData = new List<string>();

        foreach (SplitKey requested in requestedKeys)
        {
            if (spelling.TryGetValue(requested, out SplitKey? fromData))
            {
                if (!resolved.Contains(fromData)) resolved.Add(fromData);
            }
            else
            {
                notInData.Add(requested.Display);
                if (!resolved.Contains(requested)) resolved.Add(requested);
            }
        }

        if (notInData.Count == resolved.Count)
            plan.Problems.Add($"'{name}': żadna z wartości nie występuje w danych ({string.Join(", ", notInData)}) - plik byłby pusty");
        else if (notInData.Count > 0)
            plan.Warnings.Add($"'{name}': tych wartości nie ma w danych, plik powstanie bez nich: {string.Join(", ", notInData)}");

        return new Recipient(name, resolved, sheetsToKeep);
    }

    private static void ReportKeysInManyFiles(RecipientPlan plan)
    {
        var shared = plan.Files
            .SelectMany(recipient => recipient.Keys.Select(key => (Key: key, recipient.Name)))
            .GroupBy(entry => entry.Key)
            .Where(entries => entries.Count() > 1);

        foreach (var entries in shared)
            plan.Warnings.Add($"'{entries.Key.Display}' trafi do kilku plików: {string.Join(", ", entries.Select(entry => entry.Name))}");
    }

    private static void ReportValuesWithNewLine(RecipientPlan plan)
    {
        foreach (Recipient recipient in plan.Files)
            foreach (SplitKey key in recipient.Keys)
                foreach (string value in key.Values)
                    if (value.Contains('\n') || value.Contains('\r'))
                        plan.Problems.Add($"wartość '{value}' zawiera znak nowej linii - nie da się jej przekazać do filtra");
    }

    private static void ReportCollidingFileNames(SplitConfig config, RecipientPlan plan)
    {
        if (string.IsNullOrWhiteSpace(config.OutputFolder)) return;

        var collisions = plan.Files
            .GroupBy(recipient => config.OutputPathFor(recipient.Name), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1);

        foreach (var collision in collisions)
            plan.Problems.Add($"kilku odbiorców dostałoby ten sam plik {Path.GetFileName(collision.Key)}: " +
                              string.Join(", ", collision.Select(recipient => recipient.Name)));

        ReportFilesOverwritingInput(config, plan);
    }

    private static void ReportFilesOverwritingInput(SplitConfig config, RecipientPlan plan)
    {
        foreach ((string path, string what) in InputWorkbooks(config))
        {
            List<string> clashing = plan.Files
                .Where(recipient => string.Equals(
                    config.OutputPathFor(recipient.Name), path, StringComparison.OrdinalIgnoreCase))
                .Select(recipient => recipient.Name)
                .ToList();

            if (clashing.Count == 0) continue;

            plan.Problems.Add(
                $"plik odbiorcy {string.Join(", ", clashing)} nadpisałby {what} ({path}) - " +
                "zmień OutputFolder albo FileNameTemplate, żeby pliki odbiorców powstawały gdzie indziej");
        }
    }

    private static IEnumerable<(string Path, string What)> InputWorkbooks(SplitConfig config)
    {
        string? source = FullPathOrNull(() => config.SourceWorkbook);
        if (source is not null) yield return (source, "skoroszyt źródłowy");

        if (config.Mode == SplitMode.Worksheet) yield break;

        string? prepared = FullPathOrNull(config.PreparedWorkbookPath);
        if (prepared is not null) yield return (prepared, "skoroszyt przygotowany przez `prepare`");
    }

    private static string? FullPathOrNull(Func<string> path)
    {
        try
        {
            string value = path();
            return string.IsNullOrWhiteSpace(value) ? null : Path.GetFullPath(value);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return null;
        }
    }
}
