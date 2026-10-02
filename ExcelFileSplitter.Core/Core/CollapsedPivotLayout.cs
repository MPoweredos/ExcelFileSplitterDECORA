using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelFileSplitter.Core;

public sealed record CollapsedField(
    string Name, string Caption, int Items, int Collapsed, IReadOnlyList<string> CollapsedItems)
{
    public bool AllCollapsed => Collapsed == Items;

    public override string ToString() => $"{Caption} ({Collapsed} z {Items})";
}

public sealed record CollapsedPivot(
    PivotTableRef Table, IReadOnlyList<CollapsedField> Fields, IReadOnlyList<string> HiddenTables)
{
    public string Describe() =>
        $"tabela przestawna {Table} ma zwinięte pozycje w polach: {string.Join(", ", Fields)} - pod nimi " +
        $"pamięta dane z {string.Join(", ", HiddenTables.Select(table => $"'{table}'"))}";
}

public static class CollapsedPivotLayout
{
    public static List<CollapsedPivot> Find(string packagePath, IReadOnlySet<string> narrowedTables)
    {
        using var package = ZipFile.OpenRead(packagePath);
        ZipArchiveEntry? workbook = package.GetEntry("xl/workbook.xml");
        if (workbook is null) return [];

        HashSet<string> modelConnections = PivotCacheMembers.ModelConnectionIds(package);
        Dictionary<string, (string Type, string Target)> workbookRelations =
            PivotCacheMembers.Relations(package, workbook.FullName);

        var found = new List<CollapsedPivot>();
        foreach (XElement sheet in PivotCacheMembers.Load(workbook).Descendants().Where(element => element.Name.LocalName == "sheet"))
        {
            string? relationId = sheet.Attribute(XName.Get("id", PivotCacheMembers.RelationshipsNamespace))?.Value;
            if (relationId is null || !workbookRelations.TryGetValue(relationId, out var sheetPart)) continue;

            foreach (var (type, pivotPart) in PivotCacheMembers.Relations(package, sheetPart.Target).Values)
            {
                if (type != "pivotTable" || package.GetEntry(pivotPart) is not ZipArchiveEntry pivotEntry) continue;

                string? cachePart = PivotCacheMembers.Relations(package, pivotPart).Values
                    .Where(relation => relation.Type == "pivotCacheDefinition")
                    .Select(relation => relation.Target)
                    .FirstOrDefault();
                if (cachePart is null || package.GetEntry(cachePart) is not ZipArchiveEntry cacheEntry) continue;

                XElement cache = PivotCacheMembers.Load(cacheEntry);
                if (!PivotCacheMembers.IsModelCache(cache, modelConnections)) continue;

                XElement pivot = PivotCacheMembers.Load(pivotEntry);
                var table = new PivotTableRef(pivot.Attribute("name")?.Value ?? "?", sheet.Attribute("name")?.Value ?? "?");
                CollapsedPivot? collapsed = Read(table, pivot, cache, narrowedTables);
                if (collapsed is not null) found.Add(collapsed);
            }
        }

        return found;
    }

    internal static CollapsedPivot? Read(
        PivotTableRef table, XElement pivot, XElement cache, IReadOnlySet<string> narrowedTables)
    {
        List<XElement> cacheFields = Children(cache, "cacheFields", "cacheField");
        List<XElement> pivotFields = Children(pivot, "pivotFields", "pivotField");
        List<int> rowFields = Children(pivot, "rowFields", "field")
            .Select(field => int.TryParse(field.Attribute("x")?.Value, out int index) ? index : -1)
            .ToList();

        List<string?> names = rowFields
            .Select(index => index >= 0 && index < cacheFields.Count ? cacheFields[index].Attribute("name")?.Value : null)
            .ToList();
        List<string?> tables = names.Select(TableOf).ToList();
        int levels = LevelsHidingNarrowed(tables, narrowedTables);

        var fields = new List<CollapsedField>();
        int firstLevel = -1;
        for (int level = 0; level < levels; level++)
        {
            int index = rowFields[level];
            if (names[level] is not string name || index >= pivotFields.Count) continue;

            List<XElement> items = Children(pivotFields[index], "items", "item")
                .Where(item => item.Attribute("t") is null)
                .ToList();
            List<XElement> collapsed = items.Where(IsCollapsed).ToList();
            if (collapsed.Count == 0) continue;

            if (firstLevel < 0) firstLevel = level;
            Dictionary<int, string> uniqueNames = UniqueNames(cacheFields[index], name);
            fields.Add(new CollapsedField(
                name,
                CaptionOf(pivotFields[index], cacheFields[index], name),
                items.Count,
                collapsed.Count,
                collapsed
                    .Select(item => int.TryParse(item.Attribute("x")?.Value, out int shared)
                        && uniqueNames.TryGetValue(shared, out string? uniqueName) ? uniqueName : null)
                    .OfType<string>()
                    .ToList()));
        }

        return fields.Count == 0
            ? null
            : new CollapsedPivot(table, fields, HiddenNarrowedTables(tables, firstLevel, narrowedTables));
    }

    public static string? TableOf(string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(fieldName)) return null;

        List<string> parts = PivotCacheMembers.BracketParts(fieldName.Trim());
        if (parts.Count < 2 || parts[0].Equals("Measures", StringComparison.OrdinalIgnoreCase)) return null;
        return parts[0];
    }

    public static int LevelsHidingNarrowed(IReadOnlyList<string?> rowFieldTables, IReadOnlySet<string> narrowedTables)
    {
        for (int level = rowFieldTables.Count - 1; level > 0; level--)
            if (rowFieldTables[level] is { } table && narrowedTables.Contains(table)) return level;
        return 0;
    }

    public static List<string> HiddenNarrowedTables(
        IReadOnlyList<string?> rowFieldTables, int belowLevel, IReadOnlySet<string> narrowedTables) =>
        rowFieldTables.Skip(belowLevel + 1)
            .OfType<string>()
            .Where(narrowedTables.Contains)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static bool IsCollapsed(XElement item) => item.Attribute("e")?.Value is "0" or "false";

    private static Dictionary<int, string> UniqueNames(XElement cacheField, string fieldName)
    {
        var names = new Dictionary<int, string>();
        foreach (XElement cached in cacheField.Descendants().Where(element => element.Name.LocalName == "cachedUniqueName"))
        {
            if (int.TryParse(cached.Attribute("index")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
                && cached.Attribute("name")?.Value is string name)
                names.TryAdd(index, name);
        }

        List<string> parts = PivotCacheMembers.BracketParts(fieldName);
        if (parts.Count < 2) return names;

        string hierarchy = $"[{parts[0].Replace("]", "]]")}].[{parts[1].Replace("]", "]]")}].";
        List<XElement> shared = Children(cacheField, "sharedItems", null);
        for (int index = 0; index < shared.Count; index++)
        {
            string? value = shared[index].Attribute("v")?.Value;
            if (value is not null && value.StartsWith(hierarchy, StringComparison.OrdinalIgnoreCase))
                names.TryAdd(index, value);
        }

        return names;
    }

    private static string CaptionOf(XElement pivotField, XElement cacheField, string fieldName) =>
        pivotField.Attribute("name")?.Value
        ?? cacheField.Attribute("caption")?.Value
        ?? PivotCacheMembers.BracketParts(fieldName).LastOrDefault()
        ?? fieldName;

    private static List<XElement> Children(XElement parent, string container, string? child) =>
        parent.Elements()
            .Where(element => element.Name.LocalName == container)
            .Elements()
            .Where(element => child is null || element.Name.LocalName == child)
            .ToList();
}
