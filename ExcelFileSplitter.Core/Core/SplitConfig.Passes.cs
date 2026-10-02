namespace ExcelFileSplitter.Core;

public sealed partial class SplitConfig
{
    public IReadOnlyList<SplitPass> SplitPasses() => Passes.Count > 0
        ? Passes
        : [new SplitPass
        {
            SplitColumn = SplitColumn,
            SplitColumns = SplitColumns,
            Values = Values,
            Keys = Keys,
            Groups = Groups,
        }];

    public SplitConfig ForPass(SplitPass pass)
    {
        var copy = (SplitConfig)MemberwiseClone();

        copy.Passes = [];
        copy.SplitColumn = pass.SplitColumn;
        copy.SplitColumns = pass.SplitColumns;
        copy.Values = pass.Values;
        copy.Keys = pass.Keys;
        copy.Groups = pass.Groups;

        copy.SheetsToKeep = Recipients.MergeSheets(SheetsToKeep, pass.SheetsToKeep);

        if (pass.OutputFolder.Trim().Length > 0) copy.OutputFolder = pass.OutputFolder;
        copy.FileNameTemplate = ResolvePassMarker(
            pass.FileNameTemplate.Trim().Length > 0 ? pass.FileNameTemplate : FileNameTemplate, pass.Name);

        return copy;
    }

    internal static string ResolvePassMarker(string template, string passName)
    {
        string safe = SafeFileNamePart(passName);

        foreach (string marker in new[] { "{podzial}", "{pass}" })
        {
            template = safe.Length > 0
                ? template.Replace(marker, safe)
                : template.Replace("_" + marker, "").Replace("-" + marker, "")
                          .Replace(marker + "_", "").Replace(marker + "-", "")
                          .Replace(marker, "");
        }
        return template;
    }
}
