using System.Text.RegularExpressions;

namespace ExcelFileSplitter.Core;

public sealed record PivotTableRef(string Name, string Sheet)
{
    public override string ToString() => $"'{Name}' (arkusz '{Sheet}')";
}

public sealed record PivotRefreshFailure(IReadOnlyList<PivotTableRef> Tables, string Reason, string? Source)
{
    public string Explain() => $"Excel: {Reason}" + (Source is null ? "" : $"; źródło danych: {Source}");

    public (bool IsWarning, string Message) Report(IReadOnlySet<string> sheetsInFile)
    {
        if (Tables.Count == 0)
            return (true, "cache tabel przestawnych, któremu nie da się przypisać żadnej tabeli, nie dał się " +
                          "odświeżyć - trzyma dane sprzed zawężenia, a Excel może go zapisać w pliku " +
                          $"(np. dla fragmentatora). {Explain()}");

        List<PivotTableRef> staying = Tables.Where(table => sheetsInFile.Contains(table.Sheet)).ToList();
        if (staying.Count > 0)
            return (true, staying.Count == 1
                ? $"tabela przestawna {staying[0]} nie dała się odświeżyć i zostaje w pliku - pokazuje dane " +
                  $"sprzed zawężenia, w tym być może innych odbiorców. {Explain()}"
                : $"tabele przestawne {string.Join(", ", staying)} nie dały się odświeżyć i zostają w pliku - " +
                  $"pokazują dane sprzed zawężenia, w tym być może innych odbiorców. {Explain()}");

        if (Tables.Count == 1)
            return (false, $"tabela przestawna {Tables[0]} nie dała się odświeżyć, ale jej arkusz nie trafia " +
                           $"do odbiorcy. {Explain()}");

        bool oneSheet = Tables.Select(table => table.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1;
        return (false, $"tabele przestawne {string.Join(", ", Tables)} nie dały się odświeżyć, ale " +
                       (oneSheet ? "ich arkusz nie trafia" : "ich arkusze nie trafiają") +
                       $" do odbiorcy. {Explain()}");
    }
}

public sealed record PivotRefresh(
    DateTime Started, int Refreshed, IReadOnlyList<PivotRefreshFailure> Failures, IReadOnlySet<DateTime> FailedDates)
{
    private static readonly TimeSpan SameMoment = TimeSpan.FromMilliseconds(10);

    public bool IsStale(DateTime? refreshedAt) =>
        refreshedAt is null
            ? Failures.Count > 0
            : refreshedAt.Value < RefreshClock.Threshold(Started)
              || FailedDates.Any(failed => (failed - refreshedAt.Value).Duration() < SameMoment);
}

public static class PivotSource
{
    private static readonly Regex R1C1 = new(
        @"^(?<sheet>.+)!R(?<r1>\d{1,7})C(?<c1>\d{1,5})(:R(?<r2>\d{1,7})C(?<c2>\d{1,5}))?$", RegexOptions.Compiled);

    public static string? Describe(object? sourceData)
    {
        if (sourceData is not string text || string.IsNullOrWhiteSpace(text)) return null;

        Match match = R1C1.Match(text.Trim());
        if (!match.Success) return text.Trim();

        int firstRow = int.Parse(match.Groups["r1"].Value);
        int firstColumn = int.Parse(match.Groups["c1"].Value);
        int lastRow = match.Groups["r2"].Success ? int.Parse(match.Groups["r2"].Value) : firstRow;
        int lastColumn = match.Groups["c2"].Success ? int.Parse(match.Groups["c2"].Value) : firstColumn;

        string range = $"{match.Groups["sheet"].Value}!{ColumnLetters(firstColumn)}{firstRow}" +
                       (lastRow == firstRow && lastColumn == firstColumn ? "" : $":{ColumnLetters(lastColumn)}{lastRow}");
        return lastRow == firstRow ? $"{range} (sam wiersz nagłówków, bez danych)" : range;
    }

    public static string ColumnLetters(int column)
    {
        var letters = "";
        for (int rest = column; rest > 0; rest = (rest - 1) / 26)
            letters = (char)('A' + (rest - 1) % 26) + letters;
        return letters;
    }
}

public static class ExcelMessage
{
    public static string FirstSentence(string? message)
    {
        string text = (message ?? "").Trim();
        int lineEnd = text.IndexOfAny(['\r', '\n']);
        if (lineEnd >= 0) text = text[..lineEnd];

        int sentenceEnd = text.IndexOf(". ", StringComparison.Ordinal);
        if (sentenceEnd >= 0) text = text[..sentenceEnd];

        text = text.Trim().TrimEnd('.').Trim();
        return text.Length == 0 ? "bez opisu błędu" : text;
    }
}
