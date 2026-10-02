using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step2Data : UserControl, IWizardStep
{
    private readonly IWizardHost _host;
    private List<DataCandidate> _candidates = [];

    public Step2Data(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "2. Gdzie są dane";

    public Task EnterAsync()
    {
        _candidates = _host.State.Outline?.Candidates.ToList() ?? [];
        ListaKandydatow.ItemsSource = _candidates;
        ListaKandydatow.SelectedItem = Preselect();

        PoleCaleWiersze.IsChecked = _host.State.Config.Worksheet.DeleteEntireRow;
        PoleFormulyNaWartosci.IsChecked = _host.State.Config.Worksheet.FormulasToValues;

        ShowChosen();
        return Task.CompletedTask;
    }

    public void Flush()
    {
        if (ListaKandydatow.SelectedItem is not DataCandidate chosen) return;

        _host.State.Data = chosen;
        _host.State.Config.Worksheet = chosen.ToWorksheetOptions(
            PoleCaleWiersze.IsChecked == true, PoleFormulyNaWartosci.IsChecked == true);
    }

    public Task<string?> LeaveAsync()
    {
        if (ListaKandydatow.SelectedItem is not DataCandidate chosen)
            return Task.FromResult<string?>("Wskaż, gdzie są dane do podziału.");

        if (chosen.RowCount == 0)
            return Task.FromResult<string?>($"{chosen.Label} nie ma żadnych wierszy danych - wskaż inne miejsce.");

        Flush();
        return Task.FromResult<string?>(null);
    }

    private DataCandidate? Preselect()
    {
        WorksheetOptions source = _host.State.Config.Worksheet;

        if (source.Table.Length > 0)
        {
            DataCandidate? table = _candidates.FirstOrDefault(candidate =>
                string.Equals(candidate.TableName, source.Table, StringComparison.OrdinalIgnoreCase));
            if (table is not null) return table;
        }
        else if (source.Sheet.Length > 0)
        {
            DataCandidate? sheet = _candidates.FirstOrDefault(candidate =>
                candidate.TableName is null &&
                string.Equals(candidate.SheetName, source.Sheet, StringComparison.OrdinalIgnoreCase));
            if (sheet is not null) return sheet;
        }

        return _host.State.Data ?? WorkbookOutline.Best(_candidates);
    }

    private void Kandydat_Changed(object sender, SelectionChangedEventArgs e) => ShowChosen();

    private void ShowChosen()
    {
        if (ListaKandydatow.SelectedItem is not DataCandidate chosen)
        {
            PodgladNaglowkow.Text = "";
            OpisWiersza.Text = "";
            PoleWiersz.IsEnabled = false;
            PrzyciskOdczytaj.IsEnabled = false;
            return;
        }

        PoleWiersz.Text = chosen.HeaderRow.ToString(CultureInfo.InvariantCulture);
        PodgladNaglowkow.Text = chosen.Headers.Count > 0
            ? string.Join(",  ", chosen.Headers)
            : "(brak nagłówków)";

        bool rawRange = chosen.TableName is null;
        PoleWiersz.IsEnabled = rawRange;
        PrzyciskOdczytaj.IsEnabled = rawRange;
        OpisWiersza.Text = rawRange
            ? "Zmień, jeśli nagłówki nie są w pierwszym wierszu arkusza."
            : "Tabela Excela ma własny wiersz nagłówków - nie trzeba go wskazywać.";
    }

    private async void Odczytaj_Click(object sender, RoutedEventArgs e)
    {
        if (ListaKandydatow.SelectedItem is not DataCandidate chosen) return;

        if (!int.TryParse(PoleWiersz.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int headerRow)
            || headerRow < 1)
        {
            _host.Log(new ProgressLine(ProgressKind.Warning, "Wiersz nagłówków musi być liczbą większą od zera."));
            return;
        }

        var source = new WorksheetOptions { Sheet = chosen.SheetName, HeaderRow = headerRow };
        try
        {
            DataCandidate reread = await _host.RunAsync(
                $"Czytam nagłówki z wiersza {headerRow}", (workspace, _) => workspace.ReadData(source));

            int index = _candidates.IndexOf(chosen);
            if (index < 0) return;
            _candidates[index] = reread;
            ListaKandydatow.ItemsSource = null;
            ListaKandydatow.ItemsSource = _candidates;
            ListaKandydatow.SelectedItem = reread;
        }
        catch (Exception ex)
        {
            _host.Log(new ProgressLine(ProgressKind.Warning, ex.Message));
        }
    }
}
