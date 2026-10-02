using ExcelFileSplitter.Interop;

namespace ExcelFileSplitter.Core;

public sealed record PassRun(SplitPass Pass, SplitConfig Config, WorkbookInfo Info, RecipientPlan Plan)
{
    public string Label => Pass.Describe();

    public int FileCount => Plan.Files.Count;
}

public sealed record StaleOutput(string Folder, List<string> Files)
{
    public string Describe(int maxNames = 6)
    {
        string names = string.Join(", ", Files.Take(maxNames));
        if (Files.Count > maxNames) names += $" i {Files.Count - maxNames} inn.";

        return $"w folderze {Folder} leży już {Files.Count} plik(ów) Excela, które w tym przebiegu " +
               $"NIE powstaną i zostaną tam nietknięte: {names}. Sprawdź je, zanim spakujesz folder " +
               "do wysłania - mogą być z poprzedniego podziału, czyli z danymi innych odbiorców.";
    }
}

public static class SplitRun
{
    public static List<PassRun> PlanAll(
        SplitConfig config, Func<SplitConfig, WorkbookInfo?, WorkbookInfo> inspect, IProgressReporter reporter)
    {
        var runs = new List<PassRun>();
        IReadOnlyList<SplitPass> passes = config.SplitPasses();
        WorkbookInfo? previous = null;

        foreach (SplitPass pass in passes)
        {
            SplitConfig forPass = config.ForPass(pass);

            if (passes.Count > 1)
                reporter.Step($"Podział '{pass.Describe()}' po: {string.Join(" + ", forPass.SplitColumnNames)}");

            WorkbookInfo info = inspect(forPass, previous);
            previous = info;
            runs.Add(new PassRun(pass, forPass, info, Recipients.BuildPlan(forPass, info.SplitKeys)));
        }

        return runs;
    }

    public static List<string> Problems(IReadOnlyList<PassRun> runs) =>
        Labelled(runs, run => run.Plan.Problems);

    public static List<string> Warnings(IReadOnlyList<PassRun> runs) =>
        Labelled(runs, run => run.Plan.Warnings);

    public static List<string> Readiness(IReadOnlyList<PassRun> runs) =>
        Labelled(runs, run => Preflight.FindProblems(run.Info, run.Config, run.Plan.Files));

    public static int FileCount(IReadOnlyList<PassRun> runs) => runs.Sum(run => run.FileCount);

    public static List<StaleOutput> StaleOutputFiles(IReadOnlyList<PassRun> runs)
    {
        var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var folders = new List<string>();

        foreach (PassRun run in runs)
        {
            if (string.IsNullOrWhiteSpace(run.Config.OutputFolder)) continue;

            string folder = Path.GetFullPath(run.Config.OutputFolder);
            if (!folders.Contains(folder, StringComparer.OrdinalIgnoreCase)) folders.Add(folder);

            foreach (Recipient recipient in run.Plan.Files)
            {
                string path = run.Config.OutputPathFor(recipient.Name);
                planned.Add(path);

                planned.Add(SplitExecutor.RejectedPathFor(path));
            }
        }

        var stale = new List<StaleOutput>();
        foreach (string folder in folders)
        {
            List<string> found = ExcelFilesIn(folder).Where(path => !planned.Contains(path)).ToList();
            if (found.Count > 0)
                stale.Add(new StaleOutput(folder, found.Select(Path.GetFileName).OfType<string>().Order().ToList()));
        }
        return stale;
    }

    private static List<string> ExcelFilesIn(string folder)
    {
        try
        {
            if (!Directory.Exists(folder)) return [];

            return Directory
                .EnumerateFiles(folder, "*.xls*", SearchOption.TopDirectoryOnly)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .ToList();
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return [];
        }
    }

    public static List<SplitOutcome> Execute(
        ExcelSession session, IReadOnlyList<PassRun> runs, IProgressReporter reporter,
        CancellationToken cancellation = default, IReadOnlySet<string>? onlyThese = null)
    {
        var outcomes = new List<SplitOutcome>();

        foreach (PassRun run in runs)
        {
            if (cancellation.IsCancellationRequested) break;

            List<Recipient> files = Selected(run, onlyThese);
            if (files.Count == 0) continue;

            if (runs.Count > 1)
                reporter.Step($"=== Podział '{run.Label}': {files.Count} plik(ów) ===");

            using ISplitStrategy strategy = SplitExecutor.CreateStrategy(session, run.Config, run.Info);
            outcomes.AddRange(
                SplitExecutor.CreateRecipientFiles(run.Config, files, strategy, reporter, cancellation));
        }

        return outcomes;
    }

    private static List<Recipient> Selected(PassRun run, IReadOnlySet<string>? onlyThese) =>
        onlyThese is null || onlyThese.Count == 0
            ? run.Plan.Files
            : run.Plan.Files
                .Where(file => onlyThese.Contains(run.Config.OutputPathFor(file.Name)))
                .ToList();

    private static List<string> Labelled(IReadOnlyList<PassRun> runs, Func<PassRun, IEnumerable<string>> read)
    {
        var all = new List<string>();

        foreach (PassRun run in runs)
            foreach (string message in read(run))
                all.Add(runs.Count > 1 ? $"[{run.Label}] {message}" : message);

        return all;
    }
}
