namespace ExcelFileSplitter.Core;

public sealed class WorksheetOptions
{
    public string Table { get; set; } = "";

    public string Sheet { get; set; } = "";

    public int HeaderRow { get; set; } = 1;

    public bool DeleteEntireRow { get; set; } = true;

    public bool FormulasToValues { get; set; } = true;

    public bool AllowEmptyRecipients { get; set; }
}
