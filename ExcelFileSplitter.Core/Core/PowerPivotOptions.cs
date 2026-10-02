namespace ExcelFileSplitter.Core;

public sealed class PowerPivotOptions
{
    public string PreparedWorkbook { get; set; } = "";

    public string CacheFolder { get; set; } = "";

    public string ConfigSheetName { get; set; } = "_Config";

    public List<QuerySpec> Queries { get; set; } = [];

    public bool UseCache { get; set; }

    public bool RefreshPivotCaches { get; set; }

    public bool CacheSelfCheck { get; set; } = true;

    public string ParamsTable { get; set; } = "Params";

    public string SplitParamName { get; set; } = "SplitValue";
    public string UseCacheParamName { get; set; } = "UseCache";
    public string CachePathParamName { get; set; } = "CachePath";

    public string AllValuesToken { get; set; } = "__ALL__";

    public string SplitTable { get; set; } = "";

    public List<string> CacheTables { get; set; } = [];

    public HashSet<string> NarrowedTables()
    {
        var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(SplitTable)) tables.Add(SplitTable.Trim());

        foreach (QuerySpec query in Queries)
            if (query.KeyFrom is not null || !string.IsNullOrWhiteSpace(query.FilterColumn))
                tables.Add(query.Name.Trim());

        return tables;
    }
}
