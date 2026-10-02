namespace ExcelFileSplitter.Core;

public sealed class SplitKey : IEquatable<SplitKey>
{
    public SplitKey(IEnumerable<string> values) =>
        Values = values.Select(value => value.Trim()).ToArray();

    public static SplitKey Of(params string[] values) => new(values);

    public static List<SplitKey> From(IReadOnlyList<List<string>> keys, IReadOnlyList<string> values) =>
        keys.Count > 0
            ? keys.Select(key => new SplitKey(key)).ToList()
            : values.Select(value => Of(value)).ToList();

    public IReadOnlyList<string> Values { get; }

    public int Count => Values.Count;

    public bool IsIncomplete => Values.Any(value => value.Length == 0);

    public string Display => string.Join(" - ", Values);

    public bool Equals(SplitKey? other)
    {
        if (other is null || other.Values.Count != Values.Count) return false;

        for (int i = 0; i < Values.Count; i++)
            if (!string.Equals(Values[i], other.Values[i], StringComparison.OrdinalIgnoreCase))
                return false;

        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as SplitKey);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (string value in Values) hash.Add(value, StringComparer.OrdinalIgnoreCase);
        return hash.ToHashCode();
    }

    public override string ToString() => Display;
}

public interface IKeySelection
{
    List<string> Values { get; }

    List<List<string>> Keys { get; }
}

public static class KeySelection
{
    public static List<SplitKey> RequestedKeys(this IKeySelection selection) =>
        SplitKey.From(selection.Keys, selection.Values);
}
