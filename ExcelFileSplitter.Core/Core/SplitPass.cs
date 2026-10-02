using System.Text.Json.Serialization;

namespace ExcelFileSplitter.Core;

public sealed class RecipientGroup : IKeySelection
{
    public string Name { get; set; } = "";

    public List<string> Values { get; set; } = [];

    public List<List<string>> Keys { get; set; } = [];

    public List<string> SheetsToKeep { get; set; } = [];
}
public sealed class SplitPass : IKeySelection
{
    public string Name { get; set; } = "";

    public string SplitColumn { get; set; } = "";

    public List<string> SplitColumns { get; set; } = [];

    public List<string> Values { get; set; } = [];

    public List<List<string>> Keys { get; set; } = [];

    public List<RecipientGroup> Groups { get; set; } = [];

    public List<string> SheetsToKeep { get; set; } = [];

    public string FileNameTemplate { get; set; } = "";

    public string OutputFolder { get; set; } = "";

    [JsonIgnore]
    public IReadOnlyList<string> SplitColumnNames => SplitConfig.ColumnNames(SplitColumns, SplitColumn);

    public string Describe() =>
        Name.Trim().Length > 0 ? Name.Trim() : string.Join(" + ", SplitColumnNames);
}
