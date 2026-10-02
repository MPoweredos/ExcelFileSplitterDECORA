namespace ExcelFileSplitter.Core;

public sealed partial class SplitConfig
{
    public string PreparedWorkbookPath()
    {
        if (!string.IsNullOrWhiteSpace(PowerPivot.PreparedWorkbook)) return Path.GetFullPath(PowerPivot.PreparedWorkbook);
        string sourcePath = Path.GetFullPath(SourceWorkbook);
        string directory = Path.GetDirectoryName(sourcePath)!;
        return Path.Combine(directory, Path.GetFileNameWithoutExtension(sourcePath) + ".prepared" + Path.GetExtension(sourcePath));
    }

    public string WorkbookToProcess()
    {
        if (Mode == SplitMode.Worksheet) return Path.GetFullPath(SourceWorkbook);

        string preparedPath = PreparedWorkbookPath();
        return File.Exists(preparedPath) ? preparedPath : Path.GetFullPath(SourceWorkbook);
    }

    internal static string SafeFileNamePart(string text) =>
        string.Join("_", text.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim();

    public string OutputFileNameFor(string recipientName)
    {
        string safeName = SafeFileNamePart(recipientName);
        string fileName = FileNameTemplate.Replace("{name}", safeName).Replace("{value}", safeName);

        string templateExtension = Path.GetExtension(fileName);
        if (ExcelFileTypes.IsWorkbook(fileName)) fileName = fileName[..^templateExtension.Length];
        return fileName + Path.GetExtension(SourceWorkbook);
    }

    public string OutputPathFor(string recipientName) =>
        Path.Combine(Path.GetFullPath(OutputFolder), OutputFileNameFor(recipientName));
}
