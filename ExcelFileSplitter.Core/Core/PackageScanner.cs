using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelFileSplitter.Core;

public sealed record ScanHit(string PartName, string Value, bool IsQualifiedReference)
{
    public string? Context { get; init; }

    public string Describe() => Context is null ? $"{PartName} -> {Value}" : $"{PartName} -> {Value} ({Context})";
}

public static class PackageScanner
{
    private const int ContextChars = 40;
    private const int MaxTextLength = 32_767;
    private const char LastLatinLetter = '\u017F';

    private enum PartKind { Markup, Records, OtherBinary }

    public static List<ScanHit> FindForbiddenValues(
        string packagePath, IEnumerable<string> forbidden, string table, IReadOnlyList<string> columns)
    {
        var forbiddenValues = forbidden
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var hits = new List<ScanHit>();
        if (forbiddenValues.Length == 0) return hits;

        string[] columnRefPrefixes = columns.Select(column => $"[{table}].[{column}].").ToArray();
        var qualifiedForms = forbiddenValues.ToDictionary(
            value => value,
            value => columnRefPrefixes.SelectMany(prefix => new[]
            {
                prefix + "&[" + value + "]",
                prefix + "&amp;[" + value + "]",
                prefix + "[" + value + "]",
            }).ToArray(),
            StringComparer.OrdinalIgnoreCase);

        using var package = ZipFile.OpenRead(packagePath);
        foreach (ZipArchiveEntry part in package.Entries)
        {
            if (part.Length == 0 || IsImage(part.FullName)) continue;

            byte[] partBytes;
            try
            {
                using var stream = part.Open();
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                partBytes = buffer.ToArray();
            }
            catch { continue; }

            PartKind kind = KindOf(part.FullName, partBytes);
            string[] texts = Decode(partBytes, kind);

            foreach (string value in forbiddenValues)
            {
                string? context = FindIn(texts, value, kind, out bool found);
                if (!found) continue;

                bool isQualified = qualifiedForms[value]
                    .Any(form => texts.Any(text => text.Contains(form, StringComparison.OrdinalIgnoreCase)));
                hits.Add(new ScanHit(part.FullName, value, isQualified) { Context = context });
            }
        }

        return hits;
    }

    public static List<string> FindStalePivotCaches(string packagePath, Func<DateTime?, bool> isStale)
    {
        var stale = new List<string>();

        using var package = ZipFile.OpenRead(packagePath);
        foreach (ZipArchiveEntry part in package.Entries)
        {
            if (!IsPivotCacheDefinition(part.FullName)) continue;

            DateTime? refreshedAt;
            string? sourceType;
            try
            {
                using Stream stream = part.Open();
                (refreshedAt, sourceType) = ReadPivotCacheDefinition(stream);
            }
            catch (Exception exception) when (exception is XmlException or InvalidDataException or IOException)
            {
                stale.Add($"{part.FullName} - nie da się odczytać ({exception.Message})");
                continue;
            }

            if (string.Equals(sourceType, "external", StringComparison.OrdinalIgnoreCase)) continue;
            if (!isStale(refreshedAt)) continue;

            stale.Add(refreshedAt is null
                ? $"{part.FullName} - brak daty odświeżenia"
                : $"{part.FullName} - cache z {refreshedAt.Value:yyyy-MM-dd HH\\:mm\\:ss}");
        }

        return stale;
    }

    internal static bool IsPivotCacheDefinition(string partName) =>
        Path.GetFileName(partName).StartsWith("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase)
        && Path.GetExtension(partName).Equals(".xml", StringComparison.OrdinalIgnoreCase);

    private static (DateTime? RefreshedAt, string? SourceType) ReadPivotCacheDefinition(Stream stream)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true };
        using var reader = XmlReader.Create(stream, settings);

        reader.MoveToContent();
        if (reader.LocalName != "pivotCacheDefinition")
            throw new XmlException($"zamiast definicji cache'u jest element <{reader.LocalName}>");

        DateTime? refreshedAt = ParseRefreshDate(reader.GetAttribute("refreshedDate"), reader.GetAttribute("refreshedDateIso"));

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "cacheSource")
                return (refreshedAt, reader.GetAttribute("type"));
        }

        return (refreshedAt, null);
    }

    private static DateTime? ParseRefreshDate(string? oleDate, string? isoDate)
    {
        if (double.TryParse(oleDate, NumberStyles.Float, CultureInfo.InvariantCulture, out double days)
            && days is > -657_435d and < 2_958_466d)
            return DateTime.FromOADate(days);

        if (DateTime.TryParse(isoDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
            return parsed;

        return null;
    }

    private static string? FindIn(string[] texts, string value, PartKind kind, out bool found)
    {
        foreach (string text in texts)
        {
            foreach (string form in Forms(value, kind == PartKind.Markup))
            {
                int at = kind == PartKind.OtherBinary
                    ? text.IndexOf(form, StringComparison.OrdinalIgnoreCase)
                    : IndexOfWhole(text, form, records: kind == PartKind.Records);
                if (at < 0) continue;

                found = true;
                return Around(text, at, form.Length, markup: kind == PartKind.Markup);
            }
        }

        found = false;
        return null;
    }

    private static PartKind KindOf(string partName, byte[] bytes) =>
        IsMarkup(partName) ? PartKind.Markup
        : IsRecordStream(bytes) ? PartKind.Records
        : PartKind.OtherBinary;

    private static bool IsMarkup(string partName) =>
        Path.GetExtension(partName).ToLowerInvariant() is ".xml" or ".rels" or ".vml";

    private static bool IsImage(string partName) =>
        Path.GetExtension(partName).ToLowerInvariant()
            is ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".tiff" or ".emf" or ".wmf" or ".ico";

    private static bool IsRecordStream(byte[] bytes)
    {
        int position = 0;
        while (position < bytes.Length)
        {
            if (!TryReadVarInt(bytes, ref position, maxBytes: 2, out _)) return false;
            if (!TryReadVarInt(bytes, ref position, maxBytes: 4, out int size) || size > bytes.Length - position)
                return false;
            position += size;
        }
        return true;
    }

    private static bool TryReadVarInt(byte[] bytes, ref int position, int maxBytes, out int value)
    {
        value = 0;
        for (int index = 0; index < maxBytes && position < bytes.Length; index++)
        {
            byte current = bytes[position++];
            value |= (current & 0x7F) << (7 * index);
            if ((current & 0x80) == 0) return true;
        }
        return false;
    }

    private static string[] Decode(byte[] bytes, PartKind kind)
    {
        string fromSecondByte = bytes.Length > 1 ? Encoding.Unicode.GetString(bytes, 1, bytes.Length - 1) : "";
        return kind switch
        {
            PartKind.Markup => [Encoding.UTF8.GetString(bytes), Encoding.Unicode.GetString(bytes)],
            PartKind.Records => [Encoding.Unicode.GetString(bytes), fromSecondByte],
            _ => [Encoding.UTF8.GetString(bytes), Encoding.Unicode.GetString(bytes), fromSecondByte],
        };
    }

    private static IEnumerable<string> Forms(string value, bool markup)
    {
        yield return value;
        if (!markup) yield break;

        string escaped = value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        string withQuotes = escaped.Replace("\"", "&quot;");
        string withApostrophes = withQuotes.Replace("'", "&apos;");

        foreach (string form in new[] { escaped, withQuotes, withApostrophes }.Distinct(StringComparer.Ordinal))
            if (form != value) yield return form;
    }

    private static int IndexOfWhole(string text, string value, bool records)
    {
        bool checkStart = char.IsLetterOrDigit(value[0]);
        bool checkEnd = char.IsLetterOrDigit(value[^1]);

        int from = 0;
        while (from <= text.Length - value.Length)
        {
            int at = text.IndexOf(value, from, StringComparison.OrdinalIgnoreCase);
            if (at < 0) return -1;

            int end = at + value.Length;
            bool startIsEdge = !checkStart || at == 0 || !IsWordChar(text[at - 1], records)
                               || (records && TextLengthAt(text, at) >= value.Length);
            bool endIsEdge = !checkEnd || end == text.Length || !IsWordChar(text[end], records)
                             || (records && TextEndsAt(text, at, end));
            if (startIsEdge && endIsEdge) return at;

            from = at + 1;
        }
        return -1;
    }

    private static bool IsWordChar(char character, bool records) =>
        char.IsLetterOrDigit(character) && (!records || character <= LastLatinLetter);

    private static bool TextEndsAt(string text, int at, int end)
    {
        for (int start = at; start > 0 && end - start <= MaxTextLength; start--)
        {
            if (TextLengthAt(text, start) == end - start) return true;
            if (!IsTextChar(text[start - 1])) return false;
        }
        return false;
    }

    private static int? TextLengthAt(string text, int start)
    {
        if (start >= 2 && text[start - 1] == '\0') return text[start - 2];
        if (start >= 2 && (text[start - 2] & 0xFF00) == 0x1700) return text[start - 1];
        if (start >= 1 && text[start - 1] < ' ') return text[start - 1];
        return null;
    }

    private static bool IsTextChar(char character) =>
        !char.IsControl(character) || character is '\t' or '\n' or '\r';

    private static bool IsReadable(char character) =>
        !char.IsControl(character)
        && (character <= LastLatinLetter || character is >= '\u2010' and <= '\u2027' or '\u20AC');

    private static string Around(string text, int at, int length, bool markup)
    {
        int start = Math.Max(0, at - ContextChars);
        int end = Math.Min(text.Length, at + length + ContextChars);

        var snippet = new StringBuilder();
        if (start > 0) snippet.Append('…');
        foreach (char character in text.AsSpan(start, end - start))
        {
            if (markup)
            {
                snippet.Append(char.IsControl(character) ? ' ' : character);
                continue;
            }

            if (IsReadable(character)) snippet.Append(character);
            else if (snippet.Length == 0 || snippet[^1] != '·') snippet.Append('·');
        }
        if (end < text.Length) snippet.Append('…');
        return snippet.ToString();
    }
}
