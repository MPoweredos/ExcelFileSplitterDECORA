using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step5Run : UserControl, IWizardStep
{
    private const string Blocking = "#B00020";
    private const string Advisory = "#8A5300";

    private static readonly Brush AdvisoryBrush = new SolidColorBrush(Color.FromRgb(0x8A, 0x53, 0x00));

    private readonly IWizardHost _host;

    private List<PassRun> _runs = [];

    private int _planned;

    private List<PlanRow> _plan = [];

    private bool _readyToStart;

    public Step5Run(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
        ListaLogu.ItemsSource = _host.LogLines;
    }

    public string Title => "5. Podsumowanie i podział";

    public async Task EnterAsync()
    {
        Wynik.Text = "";
        PrzyciskStart.IsEnabled = false;
        StanGotowosci.Text = "Sprawdzam gotowość...";
        StanGotowosci.Foreground = Brushes.Black;
        RamkaProblemow.Visibility = Visibility.Collapsed;
        StarePliki.Visibility = Visibility.Collapsed;
        ListaPlanu.ItemsSource = null;
        _runs = [];
        _planned = 0;

        SplitConfig config = _host.State.Config;

        var blockers = new List<string>();

        var findings = new List<string>();

        try { config.Validate(); }
        catch (Exception ex) { blockers.Add(ex.Message); }

        try
        {
            _runs = await _host.RunAsync(
                "Sprawdzam skoroszyt przed podziałem",
                (workspace, reporter) => workspace.InspectAll(config, reporter));
        }
        catch (Exception ex)
        {
            _runs = [];
            blockers.Add($"nie udało się sprawdzić skoroszytu: {ex.Message}");
        }

        if (_runs.Count > 0)
        {
            _planned = SplitRun.FileCount(_runs);

            findings.AddRange(SplitRun.Readiness(_runs));
            findings.AddRange(SplitRun.Problems(_runs));
            foreach (string warning in SplitRun.Warnings(_runs)) _host.Log(new ProgressLine(ProgressKind.Warning, warning));

            ShowPlan();
            ShowStaleFiles();
        }

        ShowVerdict(blockers, findings);
    }

    public void Flush() { }

    public Task<string?> LeaveAsync() => Task.FromResult<string?>(null);

    private void ShowPlan()
    {
        NaglowekPlanu.Text = _runs.Count > 1
            ? $"Pliki do utworzenia ({_planned}, z {_runs.Count} podziałów)"
            : $"Pliki do utworzenia ({_planned})";

        KolumnaPodzialu.Width = _runs.Count > 1 ? 130 : 0;

        _plan = _runs
            .SelectMany(run => run.Plan.Files.Select(file =>
            {
                string path = run.Config.OutputPathFor(file.Name);
                return new PlanRow(
                    path,
                    Path.GetFileName(path),
                    DescribeOnDisk(path),
                    run.Label,
                    file.IsGroup ? string.Join(", ", file.KeyLabels) : file.Name,
                    file.SheetsToKeep.Count == 0 ? "(brak)" : string.Join(", ", file.SheetsToKeep));
            }))
            .ToList();

        ListaPlanu.ItemsSource = _plan;
        ShowSelectionCount();
    }

    private static string DescribeOnDisk(string path)
    {
        try
        {
            return System.IO.File.Exists(path)
                ? $"z {System.IO.File.GetLastWriteTime(path):dd.MM HH:mm}"
                : "";
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return "";
        }
    }

    private void ZaznaczWszystkie_Click(object sender, RoutedEventArgs e)
    {
        foreach (PlanRow row in _plan) row.Selected = true;
        ShowSelectionCount();
    }

    private void OdznaczIstniejace_Click(object sender, RoutedEventArgs e)
    {
        foreach (PlanRow row in _plan) row.Selected = !row.Exists;
        ShowSelectionCount();
    }

    private void Wybor_Changed(object sender, RoutedEventArgs e) => ShowSelectionCount();

    private void ShowSelectionCount()
    {
        int chosen = _plan.Count(row => row.Selected);
        int existing = _plan.Count(row => row.Exists && !row.Selected);

        PrzyciskStart.Content = chosen == _plan.Count
            ? "Rozpocznij podział"
            : $"Rozpocznij podział ({chosen} z {_plan.Count})";
        PrzyciskStart.IsEnabled = _readyToStart && chosen > 0;

        LicznikWyboru.Text = chosen == _plan.Count
            ? ""
            : $"Powstanie {chosen} z {_plan.Count} plików. Pozostałe {_plan.Count - chosen} " +
              (existing > 0
                  ? $"zostaną w folderze nietknięte (w tym {existing} już istniejących) - sprawdź ich datę, zanim je wyślesz."
                  : "nie powstaną.");
        LicznikWyboru.Foreground = chosen == _plan.Count ? Brushes.Black : Brushes.SaddleBrown;
    }

    private void ShowStaleFiles()
    {
        List<StaleOutput> stale = SplitRun.StaleOutputFiles(_runs);
        if (stale.Count == 0) return;

        StarePliki.Text = "Uwaga: " + string.Join("  ", stale.Select(folder => folder.Describe()));
        StarePliki.Visibility = Visibility.Visible;
    }

    private void ShowVerdict(List<string> blockers, List<string> findings)
    {
        bool skipped = _host.State.SkipReadinessCheck;

        ListaProblemow.ItemsSource = blockers.Select(text => new ProblemRow(text, Blocking))
            .Concat(findings.Select(text => new ProblemRow(text, skipped ? Advisory : Blocking)))
            .ToList();
        RamkaProblemow.Visibility = blockers.Count + findings.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

        _readyToStart = blockers.Count == 0 && _planned > 0 && (findings.Count == 0 || skipped);
        ShowSelectionCount();

        if (blockers.Count > 0)
        {
            StanGotowosci.Text = "Nie można zacząć podziału - popraw poniższe:";
            StanGotowosci.Foreground = Brushes.Firebrick;
        }
        else if (_planned == 0)
        {
            StanGotowosci.Text = "Nie ma czego tworzyć - wróć do kroku z grupami i wskaż, dla kogo mają powstać pliki.";
            StanGotowosci.Foreground = Brushes.Firebrick;
        }
        else if (findings.Count == 0)
        {
            StanGotowosci.Text = _runs.Count > 1
                ? $"Gotowe do podziału: {_planned} plik(ów) z {_runs.Count} podziałów."
                : $"Gotowe do podziału: {_planned} plik(ów).";
            StanGotowosci.Foreground = Brushes.DarkGreen;
        }
        else if (skipped)
        {
            StanGotowosci.Text = $"Kontrola gotowości zgłasza zastrzeżenia ({findings.Count}), ale start został " +
                                 "odblokowany na Twoje życzenie. Przeczytaj je, zanim uruchomisz podział:";
            StanGotowosci.Foreground = AdvisoryBrush;
        }
        else
        {
            StanGotowosci.Text = "Nie można zacząć podziału - popraw poniższe:";
            StanGotowosci.Foreground = Brushes.Firebrick;
        }
    }

    private async void Start_Click(object sender, RoutedEventArgs e)
    {
        if (_runs.Count == 0 || _planned == 0) return;

        HashSet<string> chosen = _plan.Where(row => row.Selected).Select(row => row.Path).ToHashSet();
        if (chosen.Count == 0) return;

        PrzyciskStart.IsEnabled = false;
        PrzyciskPrzerwij.IsEnabled = true;
        Wynik.Text = "";
        _host.LogLines.Clear();

        if (chosen.Count < _plan.Count)
            _host.Log(new ProgressLine(ProgressKind.Warning,
                $"powstanie {chosen.Count} z {_plan.Count} plików - pozostałe zostają w folderze " +
                "nietknięte i NIE są sprawdzane w tym przebiegu"));

        List<PassRun> runs = _runs;
        int planned = chosen.Count;

        try
        {
            List<SplitOutcome> outcomes = await _host.RunCancellableAsync(
                $"Dzielę plik na {planned} części",
                (workspace, reporter, token) => workspace.Split(runs, reporter, token, chosen));

            ShowOutcomes(outcomes, planned);
            ShowPlan();
        }
        catch (Exception ex)
        {
            Wynik.Text = $"Podział przerwany błędem: {ex.Message}";
            Wynik.Foreground = Brushes.Firebrick;
            _host.Log(new ProgressLine(ProgressKind.Warning, ex.Message));
        }
        finally
        {
            PrzyciskPrzerwij.IsEnabled = false;
            ShowSelectionCount();
        }
    }

    private void ShowOutcomes(List<SplitOutcome> outcomes, int planned)
    {
        List<SplitOutcome> rejected = outcomes.Where(outcome => !outcome.Ok).ToList();

        foreach (SplitOutcome outcome in rejected)
        {
            _host.Log(new ProgressLine(ProgressKind.Warning,
                $"{outcome.FileName}: {outcome.Error ?? outcome.Verification?.Summary() ?? "kontrola nie przeszła"}"));
        }

        if (rejected.Count > 0)
        {
            Wynik.Text = $"Gotowych plików: {outcomes.Count - rejected.Count} z {planned}. " +
                         $"{rejected.Count} NIE przeszło kontroli - mają w nazwie _NIE_WYSYŁAĆ i nie wolno ich wysyłać.";
            Wynik.Foreground = Brushes.Firebrick;
            return;
        }

        Wynik.Text = outcomes.Count < planned
            ? $"Przerwano. Gotowych i sprawdzonych plików: {outcomes.Count} z {planned}."
            : $"Gotowe: {outcomes.Count} plik(ów), każdy sprawdzony.";
        Wynik.Foreground = Brushes.DarkGreen;
    }

    private void Przerwij_Click(object sender, RoutedEventArgs e)
    {
        _host.Cancel();
        PrzyciskPrzerwij.IsEnabled = false;
        _host.Log(new ProgressLine(ProgressKind.Warning, "Przerywam po skończeniu bieżącego pliku..."));
    }

    public sealed class PlanRow(string path, string file, string onDisk, string pass, string values, string sheets)
        : INotifyPropertyChanged
    {
        private bool _selected = true;

        public string Path { get; } = path;
        public string File { get; } = file;

        public string OnDisk { get; } = onDisk;

        public string Pass { get; } = pass;
        public string Values { get; } = values;
        public string Sheets { get; } = sheets;

        public bool Selected
        {
            get => _selected;
            set
            {
                if (_selected == value) return;
                _selected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Selected)));
            }
        }

        public bool Exists => OnDisk.Length > 0;

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public sealed record ProblemRow(string Text, string Color);
}
