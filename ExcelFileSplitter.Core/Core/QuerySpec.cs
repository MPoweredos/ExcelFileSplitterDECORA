namespace ExcelFileSplitter.Core;

public sealed class KeyFrom
{
    public string Query { get; set; } = "";

    public string Column { get; set; } = "";
}

public sealed class QuerySpec
{
    public string Name { get; set; } = "";

    public string? FilterColumn { get; set; }

    public KeyFrom? KeyFrom { get; set; }

    public bool Cache { get; set; } = true;
}
