namespace ExcelFileSplitter.Core;

public static class ExcelFileTypes
{
    public static readonly string[] Extensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];

    public static readonly string[] WithDataModel = [".xlsx", ".xlsm", ".xlsb"];

    public static bool IsWorkbook(string path) => Has(path, Extensions);

    public static bool CanHoldDataModel(string path) => Has(path, WithDataModel);

    private static bool Has(string path, string[] extensions) =>
        !string.IsNullOrWhiteSpace(path)
        && extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    public static string Listed =>
        string.Join(", ", Extensions.Take(Extensions.Length - 1).Select(extension => extension.TrimStart('.')))
        + " albo " + Extensions[^1].TrimStart('.');
}
