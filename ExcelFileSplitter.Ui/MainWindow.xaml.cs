using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Interop;
using ExcelFileSplitter.Ui.Engine;
using ExcelFileSplitter.Ui.Wizard;
using Microsoft.Win32;

namespace ExcelFileSplitter.Ui;

public partial class MainWindow : Window, IWizardHost
{
    private readonly ExcelThread _excel = new();
    private readonly WindowProgressReporter _reporter;

    private Step0Mode _start;

    private List<IWizardStep> _steps = [];

    private SplitMode? _builtFor;

    private int _index;
    private bool _busy;

    private CancellationTokenSource? _cancellation;

    private bool _closeWhenIdle;

    private bool _startOverWhenIdle;

    public WizardState State { get; private set; } = new();

    public ObservableCollection<ProgressLine> LogLines { get; } = [];

    public MainWindow()
    {
        InitializeComponent();

        ShowAbout();

        _reporter = new WindowProgressReporter(Dispatcher, Log);
        _ = Task.Run(CloseExcelLeftByEarlierRuns);
        _start = new Step0Mode(this);

        StartMode(State.Config.Mode);

        _ = ShowStepAsync();
    }

    private void ShowAbout()
    {
        Assembly assembly = typeof(App).Assembly;

        string product = assembly.GetCustomAttribute<AssemblyProductAttribute>()?.Product ?? "";
        string author = assembly.GetCustomAttribute<AssemblyCompanyAttribute>()?.Company ?? "";

        string version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                         ?? assembly.GetName().Version?.ToString()
                         ?? "";
        version = version.Split('+')[0];

        if (product.Length > 0)
        {
            NazwaProgramu.Text = product;
            Title = version.Length > 0 ? $"{product}  v{version}" : product;
        }

        Autor.Text = author.Length > 0 ? $"Autor: {author}" : "";
        Wersja.Text = version.Length > 0 ? $"Wersja v{version}" : "";
    }

    public async Task<T> RunAsync<T>(string what, Func<SplitWorkspace, IProgressReporter, T> job, bool lockContent = true)
    {
        SetBusy(true, what, lockContent);
        try
        {
            return await _excel.Run(workspace => job(workspace, _reporter));
        }
        finally
        {
            SetBusy(false, "", lockContent);
        }
    }

    public async Task<T> RunCancellableAsync<T>(
        string what, Func<SplitWorkspace, IProgressReporter, CancellationToken, T> job)
    {
        _cancellation = new CancellationTokenSource();
        CancellationToken token = _cancellation.Token;

        SetBusy(true, what, lockContent: false);
        try
        {
            return await _excel.Run(workspace => job(workspace, _reporter, token));
        }
        finally
        {
            CancellationTokenSource finished = _cancellation;
            _cancellation = null;
            finished.Dispose();

            SetBusy(false, "", lockContent: false);
        }
    }

    public void Cancel() => _cancellation?.Cancel();

    public void StartMode(SplitMode mode)
    {
        State.Config.Mode = mode;

        if (mode != SplitMode.Worksheet) State.KeepOnlyFirstPass();

        if (_builtFor == mode) return;

        _builtFor = mode;
        _steps = mode == SplitMode.Worksheet
            ?
            [
                _start,
                new Step1File(this),
                new Step2Data(this),
                new Step3Column(this),
                new Step4Groups(this),
                new Step5Run(this),
            ]
            :
            [
                _start,
                new PivotStep1Files(this),
                new PivotStep2Queries(this),
                new PivotStep3Prepare(this),
                new PivotStep4Cache(this),
                new PivotStep5Column(this),
                new Step4Groups(this),
                new Step5Run(this),
            ];

        _index = Math.Min(_index, _steps.Count - 1);
    }

    private void RebuildSteps(SplitMode mode)
    {
        _builtFor = null;
        StartMode(mode);
    }

    public void GoToStep<TStep>() where TStep : IWizardStep
    {
        int index = _steps.FindIndex(step => step is TStep);
        if (index < 0) return;

        Komunikat.Text = "";
        _index = index;
        _ = ShowStepAsync();
    }

    public void Log(ProgressLine line)
    {
        LogLines.Add(line);

        if (line.Kind == ProgressKind.Warning) Komunikat.Text = line.Text;
    }

    private void SetBusy(bool busy, string what, bool lockContent)
    {
        _busy = busy;
        Pasek.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        StanPracy.Text = what;
        if (lockContent) TrescKroku.IsEnabled = !busy;
        UpdateNavigation();

        if (!busy && _closeWhenIdle) Close();
        else if (!busy && _startOverWhenIdle) Dispatcher.InvokeAsync(StartOverWhenIdle, DispatcherPriority.ApplicationIdle);
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_busy || _closeWhenIdle) return;

        bool duringSplit = _cancellation is not null;
        string question = duringSplit
            ? "Trwa podział. Zamknięcie teraz może zostawić w folderze wynikowym plik, który został " +
              "zapisany, ale jeszcze nie sprawdzony - a taki plik wygląda na gotowy do wysłania.\n\n" +
              "Przerwać podział? Bieżący plik zostanie dokończony i sprawdzony, potem okno się zamknie."
            : "Excel jest w trakcie pracy nad plikiem. Zamknąć program, gdy tylko skończy?";

        MessageBoxResult answer = MessageBox.Show(
            question, "Podział plików Excela", MessageBoxButton.YesNo, MessageBoxImage.Warning);

        e.Cancel = true;
        if (answer != MessageBoxResult.Yes) return;

        _closeWhenIdle = true;
        Cancel();
        UpdateNavigation();
        StanPracy.Text = duringSplit
            ? "Przerywam - zamknę okno, gdy bieżący plik będzie gotowy i sprawdzony."
            : "Zamknę okno, gdy Excel skończy to, co robi.";
    }

    private async void OdNowa_Click(object sender, RoutedEventArgs e)
    {
        if (_busy)
        {
            AskToStartOverWhenIdle();
            return;
        }

        bool somethingToLose = State.Config.SourceWorkbook.Length > 0 || State.ConfigPath is not null;
        if (somethingToLose)
        {
            MessageBoxResult answer = MessageBox.Show(
                "Zacząć od nowa?\n\nWszystko, co ustawiono w kreatorze, zostanie wyczyszczone, a otwarty plik - " +
                "zamknięty. Jeśli chcesz jeszcze użyć tych ustawień, najpierw kliknij „Zapisz ustawienia”.",
                "Podział plików Excela", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes) return;
        }

        await StartOverAsync();
    }

    private void AskToStartOverWhenIdle()
    {
        bool duringSplit = _cancellation is not null;
        string question = duringSplit
            ? "Trwa podział. Przerwać go i zacząć od nowa?\n\n" +
              "Bieżący plik zostanie dokończony i sprawdzony, potem kreator wróci na początek."
            : "Excel jest w trakcie pracy nad plikiem. Zacząć od nowa, gdy tylko skończy?";

        MessageBoxResult answer = MessageBox.Show(
            question, "Podział plików Excela", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes) return;

        _startOverWhenIdle = true;
        Cancel();
        UpdateNavigation();
        StanPracy.Text = duringSplit
            ? "Przerywam - zacznę od nowa, gdy bieżący plik będzie gotowy i sprawdzony."
            : "Zacznę od nowa, gdy Excel skończy to, co robi.";
    }

    private async void StartOverWhenIdle()
    {
        if (_busy || !_startOverWhenIdle || _closeWhenIdle) return;

        _startOverWhenIdle = false;
        await StartOverAsync();
    }

    private async Task StartOverAsync()
    {
        string? problem = null;
        try
        {
            await RunAsync("Zaczynam od nowa - zamykam plik i Excela", (workspace, _) =>
            {
                workspace.Reset();
                return true;
            });
        }
        catch (Exception ex)
        {
            problem = $"Nie udało się zamknąć Excela ({ex.Message.Trim()}) - kolejny plik i tak otworzy się " +
                      "w nowym procesie Excela.";
        }

        State = new WizardState();
        _start = new Step0Mode(this);
        _index = 0;
        RebuildSteps(State.Config.Mode);

        LogLines.Clear();
        Komunikat.Text = problem ?? "";
        StanPracy.Text = "Zaczynamy od nowa.";
        await ShowStepAsync();
    }

    private void UpdateNavigation()
    {
        PrzyciskWstecz.IsEnabled = !_busy && _index > 0;
        PrzyciskDalej.IsEnabled = !_busy && _index < _steps.Count - 1;
        PrzyciskWczytaj.IsEnabled = !_busy;
        PrzyciskZapisz.IsEnabled = !_busy;
        PrzyciskOdNowa.IsEnabled = !_startOverWhenIdle && !_closeWhenIdle;
    }

    private async Task ShowStepAsync()
    {
        IWizardStep step = _steps[_index];
        TrescKroku.Content = (UIElement)step;

        ListaKrokow.ItemsSource = _steps
            .Select((item, i) => new StepLabel(
                item.Title + (State.SkipFileReview && i > 0 && i < _steps.Count - 1 ? "   (pominięty)" : ""),
                i == _index))
            .ToList();

        UpdateNavigation();

        try
        {
            await step.EnterAsync();
        }
        catch (Exception ex)
        {
            Komunikat.Text = ex.Message;
        }
    }

    private async void Dalej_Click(object sender, RoutedEventArgs e)
    {
        Komunikat.Text = "";

        string? problem;
        try
        {
            problem = await _steps[_index].LeaveAsync();
        }
        catch (Exception ex)
        {
            Komunikat.Text = ex.Message;
            return;
        }

        if (problem is not null)
        {
            Komunikat.Text = problem;
            return;
        }

        if (_index >= _steps.Count - 1) return;

        _index++;
        await ShowStepAsync();
    }

    private async void Wstecz_Click(object sender, RoutedEventArgs e)
    {
        if (_index == 0) return;

        State.SkipFileReview = false;

        Flush();
        Komunikat.Text = "";
        _index--;
        await ShowStepAsync();
    }

    private async void Wczytaj_Click(object sender, RoutedEventArgs e)
    {
        var picker = new OpenFileDialog
        {
            Title = "Wczytaj ustawienia",
            Filter = "Ustawienia podziału (*.json)|*.json|Wszystkie pliki|*.*",
            CheckFileExists = true,
        };
        if (picker.ShowDialog() != true) return;

        SplitConfig config;
        try
        {
            config = SplitConfig.Load(picker.FileName);
        }
        catch (Exception ex)
        {
            Komunikat.Text = $"Nie udało się wczytać ustawień: {ex.Message}";
            return;
        }

        bool worksheet = config.Mode == SplitMode.Worksheet;
        LoadedConfigDialog? options = null;

        if (worksheet)
        {
            options = new LoadedConfigDialog(config, picker.FileName) { Owner = this };
            if (options.ShowDialog() != true) return;
        }

        State.Load(config, picker.FileName);
        State.SkipFileReview = options?.SkipFileReview ?? false;
        State.SkipReadinessCheck = options?.SkipReadinessCheck ?? false;

        RebuildSteps(config.Mode);
        LogLines.Clear();
        Komunikat.Text = "";

        if (!State.SkipFileReview)
        {
            StanPracy.Text = $"Wczytano ustawienia z {Path.GetFileName(picker.FileName)} - " +
                             "przejdź kreatorem, żeby je potwierdzić.";
            _index = 1;
            await ShowStepAsync();
            return;
        }

        await JumpToSummaryAsync(picker.FileName);
    }

    private async Task JumpToSummaryAsync(string configPath)
    {
        SplitConfig config = State.Config;

        try
        {
            var loaded = await RunAsync(
                $"Czytam {Path.GetFileName(config.SourceWorkbook)} według wczytanych ustawień",
                (workspace, reporter) =>
                {
                    WorkbookOutline outline = workspace.OpenWorkbook(config.SourceWorkbook, reporter);
                    DataCandidate data = workspace.ReadData(config.Worksheet);
                    List<SplitKeyCount> keys =
                        workspace.ReadSplitKeys(config.Worksheet, config.SplitPasses()[0].SplitColumnNames);
                    return (Outline: outline, Data: data, Keys: keys);
                });

            State.Outline = loaded.Outline;
            State.Data = loaded.Data;

            State.Current = 0;
            State.Keys = loaded.Keys;
        }
        catch (Exception ex)
        {
            State.SkipFileReview = false;
            Komunikat.Text = $"Nie udało się użyć wczytanych ustawień: {ex.Message}";
            StanPracy.Text = "Przejdź kreatorem i popraw to, co się nie zgadza.";
            _index = 0;
            await ShowStepAsync();
            return;
        }

        StanPracy.Text = $"Wczytano ustawienia z {Path.GetFileName(configPath)}.";
        _index = _steps.Count - 1;
        await ShowStepAsync();
    }

    private void Zapisz_Click(object sender, RoutedEventArgs e)
    {
        Flush();

        State.ApplyToConfig();

        var dialog = new SaveFileDialog
        {
            Title = "Zapisz ustawienia",
            Filter = "Ustawienia podziału (*.json)|*.json",
            FileName = Path.GetFileName(State.ConfigPath ?? "config.json"),
            AddExtension = true,
            DefaultExt = ".json",
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            State.Config.Save(dialog.FileName);
            State.ConfigPath = dialog.FileName;
            Komunikat.Text = "";
            StanPracy.Text = $"Zapisano ustawienia do {dialog.FileName}. " +
                             "Tego samego pliku używa wersja konsolowa: ExcelFileSplitter split <plik>";
        }
        catch (Exception ex)
        {
            Komunikat.Text = $"Nie udało się zapisać ustawień: {ex.Message}";
        }
    }

    private void Flush()
    {
        try
        {
            _steps[_index].Flush();
        }
        catch (Exception ex)
        {
            Komunikat.Text = ex.Message;
        }
    }

    private void Window_Closed(object sender, EventArgs e)
    {
        _excel.Dispose();
        ExcelProcessRegistry.KillOwn();
    }

    private void CloseExcelLeftByEarlierRuns()
    {
        int killed = ExcelProcessRegistry.KillLeftByEarlierRuns();
        if (killed > 0)
            _reporter.Warning($"zamknąłem {killed} proces(ów) Excela, które zostały po poprzednim uruchomieniu programu " +
                              "(widoczne tylko w Menedżerze zadań)");
    }

    public sealed record StepLabel(string Text, bool Current)
    {
        public string Weight => Current ? "Bold" : "Normal";
        public string Color => Current ? "#000000" : "#777777";
    }
}
