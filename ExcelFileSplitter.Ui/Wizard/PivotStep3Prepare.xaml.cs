using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PivotStep3Prepare : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    public PivotStep3Prepare(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
        ListaLogu.ItemsSource = _host.LogLines;
    }

    public string Title => "3. Przygotowanie";

    public Task EnterAsync()
    {
        Wynik.Text = "";
        ShowState();
        return Task.CompletedTask;
    }

    public void Flush() { }

    public Task<string?> LeaveAsync()
    {
        SplitConfig config = _host.State.Config;

        if (!File.Exists(config.PreparedWorkbookPath()))
            return Task.FromResult<string?>(
                "Nie ma jeszcze skoroszytu roboczego - kliknij \"Przygotuj skoroszyt\". " +
                "Bez niego dalsze etapy pracowałyby na nieprzygotowanym raporcie.");

        return Task.FromResult<string?>(null);
    }

    private void ShowState()
    {
        SplitConfig config = _host.State.Config;
        string prepared = config.PreparedWorkbookPath();
        bool exists = File.Exists(prepared);

        OpisZrodla.Text = $"Źródło: {config.SourceWorkbook}";
        OpisRoboczego.Text = $"Skoroszyt roboczy: {prepared}" +
                             (exists
                                 ? $"   (istnieje, z {File.GetLastWriteTime(prepared):yyyy-MM-dd HH:mm})"
                                 : "   (jeszcze nie istnieje)");

        PrzyciskPrzygotuj.Content = exists ? "Przygotuj ponownie" : "Przygotuj skoroszyt";
    }

    private async void Przygotuj_Click(object sender, RoutedEventArgs e)
    {
        SplitConfig config = _host.State.Config;
        PrzyciskPrzygotuj.IsEnabled = false;
        Wynik.Text = "";

        try
        {
            string prepared = await _host.RunAsync(
                "Przygotowuję skoroszyt - kopiowanie i wstrzykiwanie filtrów w kod M",
                (workspace, reporter) => workspace.Prepare(config, reporter),
                lockContent: false);

            Wynik.Text = $"Gotowe: {Path.GetFileName(prepared)}";
            Wynik.Foreground = Brushes.DarkGreen;
        }
        catch (Exception ex)
        {
            Wynik.Text = $"Nie udało się przygotować skoroszytu: {ex.Message}";
            Wynik.Foreground = Brushes.Firebrick;
            _host.Log(new ProgressLine(ProgressKind.Warning, ex.Message));
        }
        finally
        {
            PrzyciskPrzygotuj.IsEnabled = true;
            ShowState();
        }
    }
}
