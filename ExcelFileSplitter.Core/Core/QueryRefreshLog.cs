using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ExcelFileSplitter.Core;

public sealed record QueryFill(string Query, DateTime? LastFilledUtc, long? Rows);

public static class QueryRefreshLog
{
    private const string SectionPrefix = "Section1/";

    public static Dictionary<string, QueryFill>? Read(string packagePath)
    {
        using var package = ZipFile.OpenRead(packagePath);
        foreach (ZipArchiveEntry entry in package.Entries)
        {
            if (!entry.FullName.StartsWith("customXml/item", StringComparison.OrdinalIgnoreCase)
                || !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;

            XElement? root = TryLoad(entry);
            if (root is null || root.Name.LocalName != "DataMashup") continue;

            return ReadMetadata(Convert.FromBase64String(root.Value.Trim()));
        }
        return null;
    }

    public static List<QueryFill> NotRefreshed(
        IReadOnlyDictionary<string, QueryFill> fills, IEnumerable<string> queries, DateTime refreshStartedUtc)
    {
        DateTime threshold = RefreshClock.Threshold(refreshStartedUtc);
        var stale = new List<QueryFill>();
        foreach (string query in queries)
        {
            QueryFill fill = fills.TryGetValue(query, out QueryFill? found) ? found : new QueryFill(query, null, null);
            if (fill.LastFilledUtc is null || fill.LastFilledUtc.Value < threshold) stale.Add(fill);
        }
        return stale;
    }

    public static Dictionary<string, QueryFill> ReadMetadata(byte[] mashup)
    {
        int position = 4;
        position += 4 + ReadLength(mashup, position);
        position += 4 + ReadLength(mashup, position);
        int metadataLength = ReadLength(mashup, position);
        position += 4;

        int xmlLength = ReadLength(mashup, position + 4);
        if (xmlLength > metadataLength - 8) throw new InvalidDataException("uszkodzone metadane zapytań Power Query");

        string xml = Encoding.UTF8.GetString(mashup, position + 8, xmlLength).TrimStart('﻿');
        return ParseItems(XElement.Parse(xml));
    }

    private static Dictionary<string, QueryFill> ParseItems(XElement metadata)
    {
        var fills = new Dictionary<string, QueryFill>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement item in metadata.Descendants().Where(element => element.Name.LocalName == "Item"))
        {
            XElement? location = item.Elements().FirstOrDefault(element => element.Name.LocalName == "ItemLocation");
            string? type = location?.Elements().FirstOrDefault(element => element.Name.LocalName == "ItemType")?.Value;
            string? path = location?.Elements().FirstOrDefault(element => element.Name.LocalName == "ItemPath")?.Value;
            if (type != "Formula" || path is null || !path.StartsWith(SectionPrefix, StringComparison.Ordinal)) continue;

            string query = Uri.UnescapeDataString(path[SectionPrefix.Length..]);
            var entries = item.Descendants()
                .Where(element => element.Name.LocalName == "Entry")
                .Select(element => (Type: element.Attribute("Type")?.Value, Value: element.Attribute("Value")?.Value))
                .Where(entry => entry.Type is not null && entry.Value is not null)
                .GroupBy(entry => entry.Type!)
                .ToDictionary(group => group.Key, group => group.First().Value!);

            fills[query] = new QueryFill(query, DateOf(entries, "FillLastUpdated"), NumberOf(entries, "FillCount"));
        }
        return fills;
    }

    private static DateTime? DateOf(Dictionary<string, string> entries, string type) =>
        entries.TryGetValue(type, out string? value) && value.StartsWith('d')
        && DateTime.TryParse(value[1..], CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime date)
            ? DateTime.SpecifyKind(date, DateTimeKind.Utc)
            : null;

    private static long? NumberOf(Dictionary<string, string> entries, string type) =>
        entries.TryGetValue(type, out string? value) && value.StartsWith('l')
        && long.TryParse(value[1..], NumberStyles.Integer, CultureInfo.InvariantCulture, out long number)
            ? number
            : null;

    private static int ReadLength(byte[] bytes, int position)
    {
        if (position < 0 || position + 4 > bytes.Length) throw new InvalidDataException("uszkodzony pakiet zapytań Power Query");
        int length = BitConverter.ToInt32(bytes, position);
        if (length < 0 || position + 4 + length > bytes.Length) throw new InvalidDataException("uszkodzony pakiet zapytań Power Query");
        return length;
    }

    private static XElement? TryLoad(ZipArchiveEntry entry)
    {
        try
        {
            using Stream stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            return XElement.Load(reader);
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
