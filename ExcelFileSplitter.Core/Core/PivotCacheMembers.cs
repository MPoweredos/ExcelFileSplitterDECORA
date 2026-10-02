using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using ExcelFileSplitter.Interop;

namespace ExcelFileSplitter.Core;

public sealed record CachedMember(string Caption, string? Key);

public sealed record CachedModelField(
    string Owner, string Field, string Table, string Column, IReadOnlyList<CachedMember> Members);

public sealed class ModelValues
{
    private readonly HashSet<string> _texts = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<double> _numbers = [];
    private readonly HashSet<DateTime> _dates = [];

    public ModelValues(IEnumerable<string> values, CultureInfo culture)
    {
        foreach (string raw in values)
        {
            string value = raw.Trim();
            _texts.Add(value);
            if (double.TryParse(value, NumberStyles.Float, culture, out double number)) _numbers.Add(number);
            if (DateTime.TryParse(value, culture, DateTimeStyles.None, out DateTime date)) _dates.Add(date);
        }
    }

    public bool Contains(CachedMember member)
    {
        string caption = member.Caption.Trim();
        string? key = member.Key?.Trim();

        if (key is { Length: 0 } || CachedData.IsSyntheticMember(caption)) return true;
        if (_texts.Contains(caption)) return true;
        if (key is null) return SlicerItems.LooksLikeNumberOrDate(caption);

        return _texts.Contains(key)
               || (double.TryParse(key, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                   && _numbers.Contains(number))
               || (DateTime.TryParse(key, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)
                   && _dates.Contains(date));
    }
}

public static class PivotCacheMembers
{
    internal const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static List<CachedModelField> Read(string packagePath)
    {
        using var package = ZipFile.OpenRead(packagePath);
        HashSet<string> modelConnections = ModelConnectionIds(package);
        Dictionary<string, List<string>> owners = PivotTablesByCache(package);

        var fields = new List<CachedModelField>();
        foreach (ZipArchiveEntry part in package.Entries)
        {
            if (!PackageScanner.IsPivotCacheDefinition(part.FullName)) continue;

            XElement definition = Load(part);
            if (!IsModelCache(definition, modelConnections)) continue;

            string owner = owners.TryGetValue(part.FullName, out List<string>? tables)
                ? string.Join(", ", tables)
                : part.FullName;

            IEnumerable<XElement> cacheFields = definition.Elements()
                .Where(element => element.Name.LocalName == "cacheFields")
                .Elements()
                .Where(element => element.Name.LocalName == "cacheField");
            foreach (XElement cacheField in cacheFields)
            {
                CachedModelField? field = ReadField(cacheField, owner);
                if (field is not null) fields.Add(field);
            }
        }
        return fields;
    }

    public static List<string> OutsideModel(CachedModelField field, ModelValues values) =>
        field.Members
            .Where(member => !values.Contains(member))
            .Select(member => member.Caption.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static CachedModelField? ReadField(XElement cacheField, string owner)
    {
        string name = cacheField.Attribute("name")?.Value ?? "";
        List<string> parts = BracketParts(name);
        if (parts.Count < 3 || parts[0].Equals("Measures", StringComparison.OrdinalIgnoreCase)) return null;

        XElement? sharedItems = cacheField.Elements().FirstOrDefault(element => element.Name.LocalName == "sharedItems");
        if (sharedItems is null) return null;

        var uniqueNames = cacheField.Descendants()
            .Where(element => element.Name.LocalName == "cachedUniqueName")
            .Select(element => (Index: element.Attribute("index")?.Value, Name: element.Attribute("name")?.Value))
            .Where(entry => entry.Index is not null && entry.Name is not null)
            .GroupBy(entry => entry.Index!)
            .ToDictionary(group => group.Key, group => group.First().Name!);

        string hierarchy = $"[{parts[0].Replace("]", "]]")}].[{parts[1].Replace("]", "]]")}]";
        var members = new List<CachedMember>();
        int index = 0;
        foreach (XElement item in sharedItems.Elements())
        {
            string? value = item.Attribute("v")?.Value;
            string? caption = item.Attribute("c")?.Value;
            string itemIndex = (index++).ToString(CultureInfo.InvariantCulture);
            if (value is null || item.Name.LocalName == "e") continue;

            string? key = uniqueNames.TryGetValue(itemIndex, out string? uniqueName)
                ? KeyOf(uniqueName, hierarchy)
                : KeyOf(value, hierarchy);
            bool valueIsUniqueName = !uniqueNames.ContainsKey(itemIndex) && key is not null;
            members.Add(new CachedMember(caption ?? (valueIsUniqueName ? key! : value), key));
        }

        if (members.Count == 0) return null;
        string label = cacheField.Attribute("caption")?.Value ?? parts[^1];
        return new CachedModelField(owner, label, parts[0], parts[^1], members);
    }

    public static string? KeyOf(string uniqueName, string hierarchy)
    {
        string prefix = hierarchy + ".&[";
        if (!uniqueName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !uniqueName.EndsWith(']')) return null;

        string escaped = uniqueName[prefix.Length..^1];
        if (escaped.Replace("]]", "").Contains(']')) return null;
        return escaped.Replace("]]", "]");
    }

    public static List<string> BracketParts(string name)
    {
        var parts = new List<string>();
        int position = 0;
        while (position < name.Length)
        {
            if (name[position] != '[') return [];

            var part = new System.Text.StringBuilder();
            position++;
            while (true)
            {
                if (position >= name.Length) return [];
                if (name[position] == ']')
                {
                    if (position + 1 < name.Length && name[position + 1] == ']')
                    {
                        part.Append(']');
                        position += 2;
                        continue;
                    }
                    position++;
                    break;
                }
                part.Append(name[position++]);
            }
            parts.Add(part.ToString());

            if (position == name.Length) break;
            if (name[position] != '.') return [];
            position++;
        }
        return parts;
    }

    internal static bool IsModelCache(XElement definition, IReadOnlySet<string> modelConnections)
    {
        XElement? source = definition.Elements().FirstOrDefault(element => element.Name.LocalName == "cacheSource");
        return source?.Attribute("type")?.Value == "external"
               && modelConnections.Contains(source.Attribute("connectionId")?.Value ?? "");
    }

    internal static HashSet<string> ModelConnectionIds(ZipArchive package)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        ZipArchiveEntry? entry = package.GetEntry("xl/connections.xml");
        if (entry is null) return ids;

        foreach (XElement connection in Load(entry).Elements().Where(element => element.Name.LocalName == "connection"))
        {
            bool model = connection.Attribute("name")?.Value == "ThisWorkbookDataModel"
                         || connection.Descendants().Any(element =>
                             element.Name.LocalName == "connection" && element.Attribute("model")?.Value is "1" or "true");
            string? id = connection.Attribute("id")?.Value;
            if (model && id is not null) ids.Add(id);
        }
        return ids;
    }

    private static Dictionary<string, List<string>> PivotTablesByCache(ZipArchive package)
    {
        var owners = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        ZipArchiveEntry? workbook = package.GetEntry("xl/workbook.xml");
        if (workbook is null) return owners;

        Dictionary<string, (string Type, string Target)> workbookRelations = Relations(package, workbook.FullName);
        foreach (XElement sheet in Load(workbook).Descendants().Where(element => element.Name.LocalName == "sheet"))
        {
            string? relationId = sheet.Attribute(XName.Get("id", RelationshipsNamespace))?.Value;
            if (relationId is null || !workbookRelations.TryGetValue(relationId, out var sheetPart)) continue;

            foreach (var (type, pivotPart) in Relations(package, sheetPart.Target).Values)
            {
                if (type != "pivotTable" || package.GetEntry(pivotPart) is not ZipArchiveEntry pivotEntry) continue;

                string pivotName = Load(pivotEntry).Attribute("name")?.Value ?? "?";
                foreach (var (cacheType, cachePart) in Relations(package, pivotPart).Values)
                {
                    if (cacheType != "pivotCacheDefinition") continue;
                    if (!owners.TryGetValue(cachePart, out List<string>? list)) owners[cachePart] = list = [];
                    list.Add(new PivotTableRef(pivotName, sheet.Attribute("name")?.Value ?? "?").ToString());
                }
            }
        }
        return owners;
    }

    internal static Dictionary<string, (string Type, string Target)> Relations(ZipArchive package, string partName)
    {
        var relations = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
        string folder = Path.GetDirectoryName(partName)?.Replace('\\', '/') ?? "";
        ZipArchiveEntry? entry = package.GetEntry($"{folder}/_rels/{Path.GetFileName(partName)}.rels".TrimStart('/'));
        if (entry is null) return relations;

        foreach (XElement relation in Load(entry).Elements())
        {
            string? id = relation.Attribute("Id")?.Value;
            string? target = relation.Attribute("Target")?.Value;
            if (id is null || target is null || relation.Attribute("TargetMode")?.Value == "External") continue;

            string type = relation.Attribute("Type")?.Value.Split('/')[^1] ?? "";
            relations[id] = (type, Resolve(folder, target));
        }
        return relations;
    }

    private static string Resolve(string folder, string target)
    {
        var segments = new List<string>();
        if (!target.StartsWith('/')) segments.AddRange(folder.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (string segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else if (segment != ".") segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    internal static XElement Load(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        return XElement.Load(reader);
    }
}
