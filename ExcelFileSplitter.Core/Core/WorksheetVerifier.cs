using System.Diagnostics;

namespace ExcelFileSplitter.Core;

public sealed record WorksheetVerificationResult(
    string FilePath,
    IReadOnlyList<SplitKey> ExpectedKeys,
    List<SplitKey> KeysInData,
    long KeptRowCount,
    bool DataSheetInFile,
    List<string> StaleSlicerItems,
    List<string> BrokenFormulas,
    List<string> Leaks,
    List<string> NameCoincidences,
    string? Error) : IVerificationResult
{
    public bool EmptyAllowed { get; init; }

    public List<string> StalePivotCaches { get; init; } = [];

    public List<SplitKey> ForeignKeys => KeysInData.Where(key => !ExpectedKeys.Contains(key)).ToList();

    public List<SplitKey> KeysWithoutData => ExpectedKeys.Where(key => !KeysInData.Contains(key)).ToList();

    public bool Ok => Error is null
                      && (KeptRowCount > 0 || EmptyAllowed)
                      && (KeysInData.Count > 0 || EmptyAllowed)
                      && ForeignKeys.Count == 0
                      && StaleSlicerItems.Count == 0
                      && StalePivotCaches.Count == 0
                      && BrokenFormulas.Count == 0
                      && Leaks.Count == 0;

    public string Summary()
    {
        if (Error is not null) return $"BŁĄD WERYFIKACJI: {Error}";

        if (Ok)
        {
            if (KeptRowCount == 0)
                return $"OK: plik pusty - dla {string.Join(", ", ExpectedKeys.Select(key => key.Display))} " +
                       "nie ma w danych ani jednego wiersza (dopuszczone ustawieniem Worksheet.AllowEmptyRecipients)";

            string scope = ExpectedKeys.Count == 1
                ? $"tylko \"{ExpectedKeys[0].Display}\""
                : $"tylko kombinacje odbiorcy ({KeysInData.Count} z {ExpectedKeys.Count})";
            var notes = new List<string>();
            if (!DataSheetInFile) notes.Add("arkusz z danymi skasowany");
            if (KeysWithoutData.Count > 0)
                notes.Add($"brak danych dla: {string.Join(", ", KeysWithoutData.Select(key => key.Display))}");
            if (NameCoincidences.Count > 0) notes.Add($"zbiegi nazw z danymi odbiorcy: {string.Join(", ", NameCoincidences.Take(3))}");
            return $"OK: {scope}, {KeptRowCount:N0} wierszy; formuły, tabele przestawne, fragmentatory i paczka czyste" +
                   (notes.Count > 0 ? $" ({string.Join("; ", notes)})" : "");
        }

        var problems = new List<string>();
        if (KeptRowCount == 0 || KeysInData.Count == 0) problems.Add("w pliku nie został żaden wiersz danych");
        if (ForeignKeys.Count > 0)
            problems.Add($"w danych zostały obce kombinacje: {string.Join(", ", ForeignKeys.Take(5).Select(key => key.Display))}");
        if (BrokenFormulas.Count > 0)
            problems.Add($"{BrokenFormulas.Count} formuł(y) z #REF! po skasowaniu wierszy: " +
                         string.Join(" | ", BrokenFormulas.Take(3)));
        if (StalePivotCaches.Count > 0)
            problems.Add($"cache tabel przestawnych sprzed podziału - może trzymać dane innych odbiorców " +
                         $"({StalePivotCaches.Count}): " + string.Join(" | ", StalePivotCaches.Take(4)));
        if (StaleSlicerItems.Count > 0)
            problems.Add($"fragmentatory pokazują {StaleSlicerItems.Count} element(ów) spoza danych: " +
                         string.Join(" | ", StaleSlicerItems.Take(4)));
        if (Leaks.Count > 0)
            problems.Add($"WYCIEK w paczce ({Leaks.Count}): " + string.Join(" | ", Leaks.Take(4)));
        return "WERYFIKACJA: " + string.Join("; ", problems);
    }
}

public static class WorksheetVerifier
{
    public static WorksheetVerificationResult VerifyRecipientFile(
        string path, SplitConfig config, Recipient recipient,
        IReadOnlyList<SplitKey> allSplitKeys, SplitEvidence? evidence, IProgressReporter reporter)
    {
        IReadOnlyList<SplitKey> expectedKeys = recipient.Keys;

        if (evidence is null)
            return new WorksheetVerificationResult(
                path, expectedKeys, [], 0, false, [], [], [], [],
                "podział nie zostawił dowodów - nie ma jak sprawdzić, czyje dane zostały w pliku");

        List<string> forbiddenValues = ForbiddenValues(allSplitKeys, expectedKeys);

        try
        {
            var stopwatch = Stopwatch.StartNew();
            List<ScanHit> hits = PackageScanner.FindForbiddenValues(
                path, forbiddenValues, config.Worksheet.Table, config.SplitColumnNames);
            List<string> staleInFile = PackageScanner.FindStalePivotCaches(path, evidence.PivotRefresh.IsStale);
            reporter.Detail($"  skan paczki {stopwatch.Elapsed:mm\\:ss} ({hits.Count} trafień)");

            (List<string> leaks, List<string> nameCoincidences) = ClassifyHits(hits, evidence.LegalInOtherColumns);

            return new WorksheetVerificationResult(
                path, expectedKeys, evidence.KeysInData, evidence.KeptRowCount, evidence.DataSheetInFile,
                evidence.StaleSlicerItems, evidence.BrokenFormulas, leaks, nameCoincidences, null)
            {
                EmptyAllowed = config.Worksheet.AllowEmptyRecipients,
                StalePivotCaches = [.. evidence.StalePivotTables, .. staleInFile],
            };
        }
        catch (Exception ex)
        {
            return new WorksheetVerificationResult(path, expectedKeys, [], 0, false, [], [], [], [], ex.Message);
        }
    }

    private static List<string> ForbiddenValues(
        IReadOnlyList<SplitKey> allSplitKeys, IReadOnlyList<SplitKey> expectedKeys)
    {
        var allowed = new HashSet<string>(
            expectedKeys.SelectMany(key => key.Values), StringComparer.OrdinalIgnoreCase);

        return allSplitKeys
            .SelectMany(key => key.Values)
            .Where(value => !allowed.Contains(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static (List<string> Leaks, List<string> NameCoincidences) ClassifyHits(
        List<ScanHit> hits, HashSet<string>? legalInOtherColumns)
    {
        var leaks = new List<string>();
        var nameCoincidences = new List<string>();

        foreach (ScanHit hit in hits)
        {
            if (legalInOtherColumns is not null && legalInOtherColumns.Contains(hit.Value))
            {
                if (!nameCoincidences.Contains(hit.Value, StringComparer.OrdinalIgnoreCase)) nameCoincidences.Add(hit.Value);
                continue;
            }

            string description = hit.Describe();
            if (legalInOtherColumns is null) description += " (nie udało się sprawdzić w danych)";
            leaks.Add(description);
        }

        return (leaks, nameCoincidences);
    }
}
