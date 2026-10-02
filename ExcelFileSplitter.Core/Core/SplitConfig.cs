using System.Text.Json.Serialization;

namespace ExcelFileSplitter.Core;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum SplitMode
{
    PowerPivot,

    Worksheet,
}

public sealed partial class SplitConfig : IKeySelection
{
    public SplitMode Mode { get; set; } = SplitMode.PowerPivot;

    public string SourceWorkbook { get; set; } = "";

    public string OutputFolder { get; set; } = "";

    public string FileNameTemplate { get; set; } = "Raport_{value}.xlsx";

    public string SplitColumn { get; set; } = "";

    public List<string> SplitColumns { get; set; } = [];

    public List<string> Values { get; set; } = [];

    public List<List<string>> Keys { get; set; } = [];

    public List<RecipientGroup> Groups { get; set; } = [];

    public List<SplitPass> Passes { get; set; } = [];

    public List<string> SheetsToKeep { get; set; } = [];

    public WorksheetOptions Worksheet { get; set; } = new();

    public PowerPivotOptions PowerPivot { get; set; } = new();

    public HardeningOptions Hardening { get; set; } = new();

    public bool Verify { get; set; } = true;

    [JsonIgnore]
    public IReadOnlyList<string> SplitColumnNames => ColumnNames(SplitColumns, SplitColumn);

    internal static IReadOnlyList<string> ColumnNames(List<string> columns, string single) =>
        columns.Count > 0
            ? columns.Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name.Trim()).ToList()
            : string.IsNullOrWhiteSpace(single) ? [] : [single.Trim()];
}
