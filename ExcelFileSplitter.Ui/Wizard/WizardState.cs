using System.IO;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Wizard;

public sealed class PassDraft
{
    public SplitPass Pass { get; init; } = new();

    public List<SplitKeyCount> Keys { get; set; } = [];

    public string Label
    {
        get
        {
            string described = Pass.Describe();
            return described.Length > 0 ? described : "(nowy podział)";
        }
    }
}

public sealed class WizardState
{
    public SplitConfig Config { get; set; } = new()
    {
        Mode = SplitMode.Worksheet,
        FileNameTemplate = "Raport_{value}.xlsx",
    };

    public WorkbookOutline? Outline { get; set; }

    public DataCandidate? Data { get; set; }

    public List<PassDraft> Drafts { get; private set; } = [new PassDraft()];

    public int Current { get; set; }

    public PassDraft CurrentDraft => Drafts[Math.Clamp(Current, 0, Drafts.Count - 1)];

    public SplitPass CurrentPass => CurrentDraft.Pass;

    public bool ManyPasses => Drafts.Count > 1;

    public List<SplitKeyCount> Keys
    {
        get => CurrentDraft.Keys;
        set => CurrentDraft.Keys = value;
    }

    public int AddPass()
    {
        Drafts.Add(new PassDraft());
        Current = Drafts.Count - 1;
        EnsurePassMarkerInTemplate();
        return Current;
    }

    private void EnsurePassMarkerInTemplate()
    {
        if (Drafts.Count < 2) return;

        string template = Config.FileNameTemplate;
        if (template.Contains("{podzial}") || template.Contains("{pass}")) return;

        if (template.Contains("{value}")) Config.FileNameTemplate = template.Replace("{value}", "{podzial}_{value}");
        else if (template.Contains("{name}")) Config.FileNameTemplate = template.Replace("{name}", "{podzial}_{name}");
    }

    public void ForgetKeys()
    {
        foreach (PassDraft draft in Drafts) draft.Keys = [];
    }

    public void KeepOnlyFirstPass()
    {
        if (Drafts.Count <= 1) return;

        Drafts = [Drafts[0]];
        Current = 0;
        ApplyToConfig();
    }

    public bool RemovePass(int index)
    {
        if (Drafts.Count <= 1 || index < 0 || index >= Drafts.Count) return false;

        Drafts.RemoveAt(index);
        Current = Math.Clamp(Current, 0, Drafts.Count - 1);
        return true;
    }

    public void ApplyToConfig()
    {
        if (Drafts.Count == 1)
        {
            SplitPass only = Drafts[0].Pass;

            Config.Passes = [];
            Config.SplitColumn = only.SplitColumn;
            Config.SplitColumns = only.SplitColumns;
            Config.Values = only.Values;
            Config.Keys = only.Keys;
            Config.Groups = only.Groups;
            Config.SheetsToKeep = only.SheetsToKeep;
            return;
        }

        Config.SplitColumn = "";
        Config.SplitColumns = [];
        Config.Values = [];
        Config.Keys = [];
        Config.Groups = [];

        Config.SheetsToKeep = [];
        Config.Passes = Drafts.Select(draft => draft.Pass).ToList();
    }

    private void RebuildDrafts()
    {
        var passes = Config.SplitPasses().ToList();

        if (Config.Passes.Count == 0 && passes.Count == 1)
            passes[0].SheetsToKeep = [.. Config.SheetsToKeep];

        Drafts = passes.Select(pass => new PassDraft { Pass = pass }).ToList();
        if (Drafts.Count == 0) Drafts = [new PassDraft()];
        Current = 0;
    }

    public string? ConfigPath { get; set; }

    public bool SkipFileReview { get; set; }

    public bool SkipReadinessCheck { get; set; }

    public void Load(SplitConfig config, string path)
    {
        Config = config;
        ConfigPath = path;
        Outline = null;
        Data = null;
        SkipFileReview = false;
        SkipReadinessCheck = false;
        RebuildDrafts();
    }

    public List<string> SheetNames => Outline?.Sheets.Select(sheet => sheet.Name).ToList() ?? [];

    public bool OutlineMatchesSource =>
        Outline is not null &&
        Config.SourceWorkbook.Length > 0 &&
        string.Equals(Outline.Path, Path.GetFullPath(Config.SourceWorkbook), StringComparison.OrdinalIgnoreCase);

}
