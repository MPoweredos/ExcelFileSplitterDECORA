using System.Diagnostics;
using System.Text;
using ExcelFileSplitter;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Interop;

Console.OutputEncoding = Encoding.UTF8;

int leftovers = ExcelProcessRegistry.KillLeftByEarlierRuns();
if (leftovers > 0)
    Console.WriteLine($"Zamknięto {leftovers} proces(ów) Excela, które zostały po poprzednim uruchomieniu programu.");

var thread = new Thread(() => Environment.ExitCode = RunCommand(args), 16 * 1024 * 1024);
thread.SetApartmentState(ApartmentState.STA);
thread.Start();
thread.Join();
ExcelProcessRegistry.KillOwn();
return Environment.ExitCode;

static int RunCommand(string[] args)
{
    if (args.Length == 0 || args[0] is "-h" or "--help" or "help")
    {
        PrintHelp();
        return 0;
    }

    string command = args[0].ToLowerInvariant();
    string configPath = args.Length > 1 ? args[1] : "config.json";

    try
    {
        if (command == "init")
        {
            if (File.Exists(configPath))
            {
                Console.Error.WriteLine($"{configPath} już istnieje - nie nadpisuję.");
                return 1;
            }

            string mode = args.Length > 2 ? args[2].ToLowerInvariant() : "powerpivot";
            SplitConfig sample = mode switch
            {
                "worksheet" => SplitConfig.CreateWorksheetSample(),
                "powerpivot" => SplitConfig.CreateSample(),
                _ => throw new InvalidOperationException($"Nieznany tryb '{mode}' - dozwolone: powerpivot, worksheet."),
            };

            sample.Save(configPath);
            Console.WriteLine($"Utworzono {configPath} (tryb {sample.Mode}). " +
                              $"Uzupełnij ścieżki i uruchom: ExcelFileSplitter inspect {configPath}");
            return 0;
        }

        var config = SplitConfig.Load(ResolveConfigPath(configPath));

        switch (command)
        {
            case "inspect": return RunInspect(config);
            case "prepare": return RunPrepare(config);
            case "cache": return RunCache(config);
            case "split": return RunSplit(config);
            default:
                Console.Error.WriteLine($"Nieznana komenda: {command}");
                PrintHelp();
                return 1;
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"BŁĄD: {ex.Message}");
        return 1;
    }
}

static string ResolveConfigPath(string path)
{
    if (File.Exists(path) || Path.IsPathRooted(path)) return path;

    for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
    {
        string candidate = Path.Combine(directory.FullName, path);
        if (File.Exists(candidate)) return candidate;
        if (directory.EnumerateFiles("*.csproj").Any()) break;
    }
    return path;
}

static WorkbookInfo ReadWorkbook(ExcelSession session, SplitConfig config, IProgressReporter reporter) =>
    config.Mode == SplitMode.Worksheet
        ? WorksheetInspector.ReadWorkbookInfo(session, config, reporter)
        : WorkbookInspector.ReadWorkbookInfo(session, config, reporter);

static int RunInspect(SplitConfig config)
{
    using var session = new ExcelSession();
    var reporter = new ConsoleProgressReporter("inspect");
    session.ReportStartupProblems(reporter);

    List<PassRun> runs = SplitRun.PlanAll(
        config, (forPass, _) => ReadWorkbook(session, forPass, reporter), reporter);
    WorkbookInfo info = runs[0].Info;

    var requiredBySheet = info.Dependencies
        .GroupBy(dependency => dependency.Sheet, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(
            group => group.Key,
            group => group.Select(dependency => dependency.Requires).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            StringComparer.OrdinalIgnoreCase);

    Console.WriteLine("\n== Arkusze (kandydaci do SheetsToKeep) ==");
    foreach (var sheet in info.Sheets)
    {
        var notes = new List<string>();
        if (!info.VisibleSheets.Contains(sheet, StringComparer.OrdinalIgnoreCase)) notes.Add("ukryty");
        if (string.Equals(sheet, config.PowerPivot.ConfigSheetName, StringComparison.OrdinalIgnoreCase))
            notes.Add("arkusz parametrów podziału - nie dodawaj go do SheetsToKeep");
        if (info.SheetsWithModelTables.TryGetValue(sheet, out var modelTableName))
            notes.Add($"trzyma tabelę modelu '{modelTableName}' - musi być w SheetsToKeep zawsze");
        if (requiredBySheet.TryGetValue(sheet, out var required))
            notes.Add("wymaga: " + string.Join(", ", required));

        Console.WriteLine(notes.Count == 0 ? $"  {sheet}" : $"  {sheet,-18} {string.Join("; ", notes)}");
    }

    if (info.Dependencies.Count > 0)
    {
        Console.WriteLine("\n== Zależności między arkuszami ==");
        foreach (var dependency in info.Dependencies)
            Console.WriteLine($"  {dependency.Sheet} -> {dependency.Requires}   ({dependency.Detail})");
    }

    Console.WriteLine(config.Mode == SplitMode.Worksheet
        ? "\n== Dane (kandydaci do SplitColumn / SplitColumns) =="
        : "\n== Tabele modelu (kandydaci do SplitTable / SplitColumn) ==");
    foreach (var table in info.ModelTables)
    {
        Console.WriteLine($"  {table.Name}  ({table.RowCount:N0} wierszy)");
        foreach (var column in table.Columns)
            Console.WriteLine($"      {column.Name}{(column.IsCalculated ? "   [obliczeniowa - liczy ją model, nie trafia do cache]" : "")}");
    }

    if (info.QueryNames.Count > 0)
    {
        Console.WriteLine("\n== Zapytania Power Query ==");
        foreach (var queryName in info.QueryNames)
            Console.WriteLine(info.OutdatedQueries.Contains(queryName) ? $"  {queryName}   [nieprzygotowane bieżącą wersją]" : $"  {queryName}");
    }

    foreach (PassRun run in runs)
    {
        if (run.Info.SplitKeys.Count == 0) continue;

        string where = run.Config.Mode == SplitMode.Worksheet
            ? string.Join(" + ", run.Config.SplitColumnNames)
            : $"{run.Config.PowerPivot.SplitTable}[{run.Config.SplitColumn}]";
        string what = run.Config.SplitColumnNames.Count > 1 ? "Kombinacje unikalne" : "Wartości unikalne";
        string whose = runs.Count > 1 ? $"podział '{run.Label}', " : "";

        Console.WriteLine($"\n== {what}: {whose}{where} ({run.Info.SplitKeys.Count}) ==");
        foreach (var key in run.Info.SplitKeys) Console.WriteLine($"  {key.Display}");
    }

    var problems = new List<string>();
    try { config.Validate(); } catch (Exception ex) { problems.Add(ex.Message); }

    problems.AddRange(SplitRun.Readiness(runs));
    problems.AddRange(SplitRun.Problems(runs));

    foreach (PassRun run in runs)
    {
        if (run.FileCount == 0 || string.IsNullOrWhiteSpace(run.Config.OutputFolder)) continue;

        string whose = runs.Count > 1 ? $" - podział '{run.Label}'" : "";
        Console.WriteLine($"\n== Plan plików ({run.FileCount}){whose} ==");

        foreach (var recipient in run.Plan.Files)
        {
            string fileName = Path.GetFileName(run.Config.OutputPathFor(recipient.Name));
            Console.WriteLine(recipient.IsGroup ? $"  {fileName}  <- {string.Join(", ", recipient.KeyLabels)}" : $"  {fileName}");
            Console.WriteLine($"      arkusze: {string.Join(", ", recipient.SheetsToKeep)}");
        }
    }
    foreach (var warning in SplitRun.Warnings(runs)) Console.WriteLine($"  uwaga: {warning}");
    foreach (var stale in SplitRun.StaleOutputFiles(runs)) Console.WriteLine($"  uwaga: {stale.Describe()}");

    Console.WriteLine(problems.Count == 0 ? "\n== Gotowość do split: OK ==" : "\n== Gotowość do split: DO POPRAWKI ==");
    foreach (var problem in problems) Console.WriteLine($"  - {problem}");

    return 0;
}

static int RunPrepare(SplitConfig config)
{
    if (RefuseInWorksheetMode(config, "prepare", "zwykły plik nie ma zapytań Power Query do przerobienia")) return 1;

    using var session = new ExcelSession();
    string preparedPath = WorkbookPreparer.PrepareWorkbook(session, config, new ConsoleProgressReporter("prepare"));
    Console.WriteLine($"\n[prepare] Gotowe: {preparedPath}");
    Console.WriteLine("[prepare] Następny krok: `cache` - jego samokontrola sprawdzi nowy skoroszyt razem z danymi.");
    return 0;
}

static bool RefuseInWorksheetMode(SplitConfig config, string command, string why)
{
    if (config.Mode != SplitMode.Worksheet) return false;
    Console.Error.WriteLine($"Komenda `{command}` dotyczy wyłącznie trybu PowerPivot - {why}.");
    Console.Error.WriteLine("W trybie Worksheet po `inspect` uruchamiasz od razu `split`.");
    return true;
}

static int RunCache(SplitConfig config)
{
    if (RefuseInWorksheetMode(config, "cache", "nie ma modelu danych do zrzucenia")) return 1;

    config.Validate();
    using var session = new ExcelSession();
    var stopwatch = Stopwatch.StartNew();
    var result = ModelCacheBuilder.BuildCache(session, config, new ConsoleProgressReporter("cache"));
    Console.WriteLine($"\n[cache] Zrzucono {result.DumpedTables.Count} tabel, " +
                      $"łącznie {result.DumpedTables.Sum(table => table.RowCount):N0} wierszy, {stopwatch.Elapsed:mm\\:ss}");

    if (!result.SelfCheckRan) return 0;
    if (result.SelfCheckPassed)
    {
        Console.WriteLine("[cache] Samokontrola OK: model zbudowany z cache ma tę samą strukturę i liczbę wierszy co oryginał.");
        return 0;
    }

    Console.Error.WriteLine("\n[cache] Samokontrola NIE przeszła - nie uruchamiaj split na tym cache:");
    foreach (var problem in result.SelfCheckProblems) Console.Error.WriteLine($"  - {problem}");
    return 1;
}

static int RunSplit(SplitConfig config)
{
    config.Validate();
    using var session = new ExcelSession();
    var reporter = new ConsoleProgressReporter("split");
    session.ReportStartupProblems(reporter);

    List<PassRun> runs = SplitRun.PlanAll(
        config, (forPass, _) => ReadWorkbook(session, forPass, reporter), reporter);

    var problems = SplitRun.Readiness(runs);
    problems.AddRange(SplitRun.Problems(runs));

    if (problems.Count > 0)
    {
        Console.Error.WriteLine("Nie zaczynam podziału:");
        foreach (var problem in problems) Console.Error.WriteLine($"  - {problem}");
        Console.Error.WriteLine("\nPopraw konfigurację i uruchom ponownie. Szczegóły pokazuje komenda `inspect`.");
        return 1;
    }

    int planned = SplitRun.FileCount(runs);
    if (planned == 0)
    {
        Console.Error.WriteLine(config.Mode == SplitMode.Worksheet
            ? "Brak plików do wygenerowania - sprawdź Worksheet, SplitColumn/SplitColumns, Values/Keys albo Groups."
            : "Brak plików do wygenerowania - sprawdź PowerPivot.SplitTable, SplitColumn, Values albo Groups.");
        return 1;
    }

    foreach (var warning in SplitRun.Warnings(runs)) reporter.Warning(warning);
    foreach (var stale in SplitRun.StaleOutputFiles(runs)) reporter.Warning(stale.Describe());

    foreach (PassRun run in runs)
    {
        string prefix = runs.Count > 1 ? $"[{run.Label}] " : "";
        reporter.Step(run.Config.Mode == SplitMode.Worksheet
            ? $"{prefix}{run.FileCount} plik(ów), tryb Worksheet, podział po: " +
              $"{string.Join(" + ", run.Config.SplitColumnNames)}, kombinacji w danych: {run.Info.SplitKeys.Count}"
            : $"{prefix}{run.FileCount} plik(ów), wartości w modelu: {run.Info.SplitKeys.Count}, " +
              $"cache: {(run.Config.PowerPivot.UseCache ? "TAK" : "nie")}");

        if (run.FileCount == 0) continue;
        reporter.Step(run.Plan.Files.Select(file => string.Join(", ", file.SheetsToKeep)).Distinct().Count() == 1
            ? $"{prefix}Arkusze u odbiorcy: {string.Join(", ", run.Plan.Files[0].SheetsToKeep)}"
            : $"{prefix}Arkusze: wspólne {(run.Config.SheetsToKeep.Count == 0 ? "(brak)" : string.Join(", ", run.Config.SheetsToKeep))}, " +
              "resztę każdy plik dokłada sam (pokazuje je `inspect`)");
    }

    var stopwatch = Stopwatch.StartNew();
    var outcomes = SplitRun.Execute(session, runs, reporter);
    stopwatch.Stop();

    var rejected = outcomes.Where(outcome => !outcome.Ok).ToList();

    Console.WriteLine($"\n[split] Zakończono: {outcomes.Count} plików w {stopwatch.Elapsed:hh\\:mm\\:ss} " +
                      $"(średnio {TimeSpan.FromTicks(outcomes.Sum(outcome => outcome.Duration.Ticks) / Math.Max(1, outcomes.Count)):mm\\:ss} na plik)");

    if (rejected.Count > 0)
    {
        Console.Error.WriteLine($"\nUWAGA: {rejected.Count} plików NIE przeszło kontroli - mają w nazwie _NIE_WYSYŁAĆ:");
        foreach (var outcome in rejected)
            Console.Error.WriteLine($"  {outcome.FileName}: {outcome.Error ?? outcome.Verification!.Summary()}");
        return 2;
    }

    return 0;
}

static void PrintHelp() => Console.WriteLine("""
    ExcelFileSplitter - dzieli raport na pliki per odbiorca, bez danych pozostałych odbiorców.

    Dwa tryby, ustawiane polem "Mode" w konfiguracji:
      PowerPivot   raport z modelem danych - filtr w kodzie M i pełne odświeżenie (prepare + cache + split)
      Worksheet    zwykły plik - dane w arkuszu albo tabeli Excela, zawężane przez skasowanie wierszy (sam split)

      init    [config.json] [powerpivot|worksheet]   tworzy przykładową konfigurację wybranego trybu
      inspect [config.json]   arkusze i zależności, dane/kolumny, plan plików, gotowość do split
      prepare [config.json]   tylko PowerPivot: kopiuje źródło i wstrzykuje parametry + filtry w kod M
      cache   [config.json]   tylko PowerPivot, ETAP 1: pełne odświeżenie ze źródeł -> CSV + samokontrola
      split   [config.json]   generuje plik per odbiorca, czyści i weryfikuje

    Kod wyjścia: 0 = OK, 1 = błąd lub konfiguracja do poprawki, 2 = pliki nie przeszły kontroli.
    Opis działania krok po kroku: docs/jak-dziala.md, tryb zwykłego pliku: docs/zwykly-excel.md.
    """);
