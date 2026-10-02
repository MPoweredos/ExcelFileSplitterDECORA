using System.Diagnostics;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public interface IVerificationResult
{
    bool Ok { get; }

    string Summary();
}

public sealed record SplitEvidence(
    List<SplitKey> KeysInData,
    long KeptRowCount,
    bool DataSheetInFile,
    HashSet<string>? LegalInOtherColumns,
    List<string> StaleSlicerItems,
    List<string> BrokenFormulas)
{
    public List<string> StalePivotTables { get; init; } = [];

    public required PivotRefresh PivotRefresh { get; init; }
}

public interface ISplitStrategy : IDisposable
{
    void PrepareOnce(IProgressReporter reporter);

    SplitEvidence? CreateRecipientWorkbook(Recipient recipient, string outputPath, IProgressReporter reporter);

    IVerificationResult Verify(string outputPath, Recipient recipient, SplitEvidence? evidence, IProgressReporter reporter);
}

public sealed record SplitOutcome(
    Recipient Recipient, string Path, TimeSpan Duration, IVerificationResult? Verification, string? Error)
{
    public bool Ok => Error is null && (Verification is null || Verification.Ok);

    public string FileName => Path.Length == 0
        ? $"(plik odbiorcy {Recipient.Name} został skasowany)"
        : System.IO.Path.GetFileName(Path);
}

public static class SplitExecutor
{
    public static List<SplitOutcome> CreateRecipientFiles(
        SplitConfig config, IReadOnlyList<Recipient> recipients, ISplitStrategy strategy, IProgressReporter reporter,
        CancellationToken cancellation = default)
    {
        Directory.CreateDirectory(Path.GetFullPath(config.OutputFolder));
        var outcomes = new List<SplitOutcome>();

        strategy.PrepareOnce(reporter);

        for (int i = 0; i < recipients.Count; i++)
        {
            if (cancellation.IsCancellationRequested)
            {
                reporter.Warning($"przerwano na żądanie użytkownika - gotowych plików: {outcomes.Count} " +
                                 $"z {recipients.Count}, reszta nie powstała");
                break;
            }

            Recipient recipient = recipients[i];
            string outputPath = config.OutputPathFor(recipient.Name);
            string rejectedPath = RejectedPathFor(outputPath);
            var stopwatch = Stopwatch.StartNew();

            if (File.Exists(rejectedPath)) TryDelete(rejectedPath, reporter);

            reporter.Step($"({i + 1}/{recipients.Count}) {recipient.Name}" +
                          (recipient.IsGroup ? $"  <- {string.Join(", ", recipient.KeyLabels)}" : ""));

            IVerificationResult? verification = null;
            string? error = null;
            try
            {
                SplitEvidence? evidence = strategy.CreateRecipientWorkbook(recipient, outputPath, reporter);
                if (config.Verify)
                {
                    var verifyTimer = Stopwatch.StartNew();
                    verification = strategy.Verify(outputPath, recipient, evidence, reporter);
                    reporter.Detail($"weryfikacja {verifyTimer.Elapsed:mm\\:ss} - {verification.Summary()}");
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                reporter.Detail($"BŁĄD: {ex.Message}");
            }

            var outcome = new SplitOutcome(recipient, outputPath, TimeSpan.Zero, verification, error);
            if (!outcome.Ok)
                outcome = outcome with
                {
                    Path = File.Exists(outputPath) ? Quarantine(outputPath, rejectedPath, reporter) : "",
                };

            stopwatch.Stop();
            reporter.Detail($"razem {stopwatch.Elapsed:mm\\:ss} -> {outcome.FileName}");
            outcomes.Add(outcome with { Duration = stopwatch.Elapsed });
        }

        return outcomes;
    }

    private static bool TryDelete(string path, IProgressReporter reporter)
    {
        try
        {
            File.Delete(path);
            return true;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            reporter.Warning($"nie udało się skasować {Path.GetFileName(path)} ({exception.Message.Trim()})");
            return false;
        }
    }

    private static string Quarantine(string outputPath, string rejectedPath, IProgressReporter reporter)
    {
        try
        {
            File.Move(outputPath, rejectedPath, overwrite: true);
            return rejectedPath;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            reporter.Warning($"nie udało się zmienić nazwy odrzuconego pliku na _NIE_WYSYŁAĆ " +
                             $"({exception.Message.Trim()}) - kasuję go, żeby nie został wysłany przez pomyłkę");
            if (TryDelete(outputPath, reporter)) return "";

            reporter.Warning($"NIE UDAŁO SIĘ SKASOWAĆ pliku {outputPath}. Zawiera on dane WSZYSTKICH " +
                             "odbiorców, a jego nazwa niczym nie ostrzega - skasuj go ręcznie, zanim " +
                             "cokolwiek wyślesz.");
            return outputPath;
        }
    }

    public static ISplitStrategy CreateStrategy(ExcelSession session, SplitConfig config, WorkbookInfo info) =>
        config.Mode == SplitMode.Worksheet
            ? new WorksheetSplit(session, config, info.SplitKeys)
            : new PowerPivotSplit(config, info.SplitKeys, info.ModelTables);

    public static string RejectedPathFor(string outputPath) =>
        Path.Combine(Path.GetDirectoryName(outputPath)!,
            Path.GetFileNameWithoutExtension(outputPath) + "_NIE_WYSYŁAĆ" + Path.GetExtension(outputPath));

    public static void DeleteSheetsNotKept(Excel.Workbook workbook, IReadOnlyList<string> sheetsToKeep)
    {
        List<string> present = Interop.Sheets.List(workbook).Select(sheet => sheet.Name).ToList();

        var missing = sheetsToKeep
            .Where(kept => !present.Contains(kept, StringComparer.OrdinalIgnoreCase))
            .ToList();
        if (missing.Count > 0)
            throw new InvalidOperationException(
                $"SheetsToKeep wskazuje arkusze, których nie ma w skoroszycie: {string.Join(", ", missing)}");

        Interop.Sheets.DeleteAllExcept(workbook, sheetsToKeep.ToList());
    }
}
