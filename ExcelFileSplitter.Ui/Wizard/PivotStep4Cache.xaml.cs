using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PivotStep4Cache : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    private bool? _fromSettings;

    public PivotStep4Cache(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
        ListaLogu.ItemsSource = _host.LogLines;
    }

    public string Title => "4. Cache (etap 1)";

    public Task EnterAsync()
    {
        Wynik.Text = "";
        _fromSettings = _host.State.Config.PowerPivot.UseCache;
        PoleUzyjCache.IsChecked = _host.State.Config.PowerPivot.UseCache;
        ShowState();
        return Task.CompletedTask;
    }

    public void Flush() => _host.State.Config.PowerPivot.UseCache = PoleUzyjCache.IsChecked == true;

    private void UzyjCache_Click(object sender, RoutedEventArgs e) => ShowState();

    public Task<string?> LeaveAsync()
    {
        Flush();
        SplitConfig config = _host.State.Config;

        if (!config.PowerPivot.UseCache) return Task.FromResult<string?>(null);

        if (config.PowerPivot.CacheFolder.Length == 0)
            return Task.FromResult<string?>(
                "Czytanie z cache wymaga folderu cache - wróć do kroku 1 i wskaż go, " +
                "albo odznacz czytanie z cache.");

        if (!HasCacheFiles(config.PowerPivot.CacheFolder))
            return Task.FromResult<string?>(
                $"W folderze {config.PowerPivot.CacheFolder} nie ma żadnych plików CSV - zbuduj cache albo " +
                "odznacz czytanie z cache, żeby podział sięgał wprost do źródeł.");

        return Task.FromResult<string?>(null);
    }

    private static bool HasCacheFiles(string folder)
    {
        try
        {
            return Directory.Exists(folder) && Directory.EnumerateFiles(folder, "*.csv").Any();
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return false;
        }
    }

    private void ShowState()
    {
        SplitConfig config = _host.State.Config;

        ShowWhichChoiceCounts();
        ShowOtherCacheSettings(config);

        if (config.PowerPivot.CacheFolder.Length == 0)
        {
            OpisFolderu.Text = "Folder cache nie jest ustawiony - wróć do kroku 1, jeśli chcesz z niego korzystać.";
            PrzyciskBuduj.IsEnabled = false;
            return;
        }

        PrzyciskBuduj.IsEnabled = true;
        int files = CountCsv(config.PowerPivot.CacheFolder);
        OpisFolderu.Text = $"Folder cache: {config.PowerPivot.CacheFolder}" +
                           (files > 0 ? $"   (jest w nim {files} plik(ów) CSV)" : "   (pusty)");
    }

    private void ShowWhichChoiceCounts()
    {
        bool chosen = PoleUzyjCache.IsChecked == true;
        string now = chosen ? "z cache" : "wprost ze źródeł";

        OpisWyboru.Text = _fromSettings switch
        {
            null => $"Liczy się ten ptaszek - podział będzie czytał {now}.",
            bool same when same == chosen =>
                $"Liczy się ten ptaszek - podział będzie czytał {now}. Tak samo mówiły wczytane ustawienia.",
            bool other =>
                $"Liczy się ten ptaszek - podział będzie czytał {now}. Wczytane ustawienia miały " +
                $"UseCache = {(other ? "tak" : "nie")}; przejście dalej nadpisze je tym wyborem.",
        };
    }

    private void ShowOtherCacheSettings(SplitConfig config)
    {
        int cachedQueries = config.PowerPivot.Queries.Count(query => query.Cache);
        string tables = config.PowerPivot.CacheTables.Count == 0
            ? "wszystkie tabele modelu"
            : string.Join(", ", config.PowerPivot.CacheTables);

        OpisPozostalych.Text =
            $"Z pliku ustawień (tego ekranu nie dotyczą): zapytania z przełącznikiem cache: {cachedQueries} " +
            $"z {config.PowerPivot.Queries.Count}; do zrzutu idzie {tables}; samokontrola po zrzucie: " +
            $"{(config.PowerPivot.CacheSelfCheck ? "tak" : "nie")}; dodatkowy przebieg po cache'ach tabel przestawnych: " +
            $"{(config.PowerPivot.RefreshPivotCaches ? "tak - to bardzo wolne" : "nie")}.";
    }

    private static int CountCsv(string folder)
    {
        try
        {
            return Directory.Exists(folder) ? Directory.EnumerateFiles(folder, "*.csv").Count() : 0;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return 0;
        }
    }

    private async void Buduj_Click(object sender, RoutedEventArgs e)
    {
        Flush();
        SplitConfig config = _host.State.Config;

        PrzyciskBuduj.IsEnabled = false;
        Wynik.Text = "";

        try
        {
            CacheResult result = await _host.RunAsync(
                "Buduję cache - pełne odświeżenie ze źródeł, to najdłuższy etap",
                (workspace, reporter) => workspace.BuildCache(config, reporter),
                lockContent: false);

            long rows = result.DumpedTables.Sum(table => table.RowCount);

            if (result.SelfCheckRan && !result.SelfCheckPassed)
            {
                Wynik.Text = $"Zrzucono {result.DumpedTables.Count} tabel, ale SAMOKONTROLA NIE PRZESZŁA - " +
                             "nie uruchamiaj podziału na tym cache:";
                Wynik.Foreground = Brushes.Firebrick;
                foreach (string problem in result.SelfCheckProblems)
                    _host.Log(new ProgressLine(ProgressKind.Warning, problem));
            }
            else
            {
                Wynik.Text = $"Gotowe: {result.DumpedTables.Count} tabel, {rows:N0} wierszy" +
                             (result.SelfCheckRan ? ", samokontrola OK" : "");
                Wynik.Foreground = Brushes.DarkGreen;
            }
        }
        catch (Exception ex)
        {
            Wynik.Text = $"Nie udało się zbudować cache: {ex.Message}";
            Wynik.Foreground = Brushes.Firebrick;
            _host.Log(new ProgressLine(ProgressKind.Warning, ex.Message));
        }
        finally
        {
            PrzyciskBuduj.IsEnabled = true;
            ShowState();
        }
    }
}
