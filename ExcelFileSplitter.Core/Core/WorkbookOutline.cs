using ExcelFileSplitter.Interop;

namespace ExcelFileSplitter.Core;

public sealed record DataCandidate(
    string? TableName,
    string SheetName,
    int HeaderRow,
    long RowCount,
    IReadOnlyList<string> Headers)
{
    public string Label => TableName is not null
        ? $"tabela '{TableName}' (arkusz {SheetName})"
        : $"arkusz '{SheetName}', nagłówki w wierszu {HeaderRow}";

    public string Summary => $"{Label} - {RowCount:N0} wierszy, {Headers.Count} kolumn";

    public WorksheetOptions ToWorksheetOptions(bool deleteEntireRow = true, bool formulasToValues = true) => new()
    {
        Table = TableName ?? "",
        Sheet = TableName is null ? SheetName : "",
        HeaderRow = HeaderRow,
        DeleteEntireRow = deleteEntireRow,
        FormulasToValues = formulasToValues,
    };

    public static DataCandidate From(DataRange range) =>
        new(range.TableName, range.SheetName, range.HeaderRow, range.RowCount, range.Headers);
}

public sealed record SplitKeyCount(SplitKey Key, long? RowCount)
{
    public string RowCountLabel => RowCount is null ? "" : $"{RowCount:N0} wierszy";

    public static List<SplitKeyCount> Unknown(IEnumerable<SplitKey> keys) =>
        keys.Select(key => new SplitKeyCount(key, null)).ToList();

    public string Display => Key.Display;

    public static List<SplitKeyCount> Tally(IEnumerable<SplitKey> keys)
    {
        var counts = new Dictionary<SplitKey, long>();

        foreach (SplitKey key in keys)
        {
            if (key.IsIncomplete) continue;
            counts[key] = counts.TryGetValue(key, out long count) ? count + 1 : 1;
        }

        return counts
            .Select(pair => new SplitKeyCount(pair.Key, pair.Value))
            .OrderBy(entry => entry.Display, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}

public sealed record WorkbookOutline(
    string Path,
    List<SheetInfo> Sheets,
    List<DataCandidate> Candidates,
    List<SheetDependency> Dependencies)
{
    public DataCandidate? BestGuess => Best(Candidates);

    public static DataCandidate? Best(IEnumerable<DataCandidate> candidates) => candidates
        .OrderByDescending(candidate => candidate.TableName is not null)
        .ThenByDescending(candidate => candidate.RowCount)
        .ThenBy(candidate => candidate.SheetName, StringComparer.OrdinalIgnoreCase)
        .FirstOrDefault();
}
