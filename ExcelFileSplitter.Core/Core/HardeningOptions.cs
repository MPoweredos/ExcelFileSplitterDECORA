namespace ExcelFileSplitter.Core;

public sealed class HardeningOptions
{
    public bool StubQueries { get; set; } = true;

    public bool DeleteQueries { get; set; }

    public bool ProtectStructure { get; set; }

    public bool DisableRefreshOnOpen { get; set; } = true;

    public bool RemoveSplitColumnSlicers { get; set; }
}
