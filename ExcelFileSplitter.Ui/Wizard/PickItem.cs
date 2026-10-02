namespace ExcelFileSplitter.Ui.Wizard;

public sealed class PickItem
{
    public required string Text { get; init; }

    public string Note { get; init; } = "";

    public bool IsSelected { get; set; }

    public object? Tag { get; init; }
}
