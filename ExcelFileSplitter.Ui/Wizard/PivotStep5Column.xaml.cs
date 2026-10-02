using System.IO;
using System.Threading.Tasks;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PivotStep5Column : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    private List<ModelTableInfo> _tables = [];

    private (string Path, DateTime WrittenUtc) _readFrom = ("", default);

    private (string Table, string Column) _loaded = ("", "");

    private bool _filling;

    public PivotStep5Column(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "5. Tabela i kolumna";

    public async Task EnterAsync()
    {
        SplitConfig config = _host.State.Config;
        (string Path, DateTime WrittenUtc) source = SourceOf(config);

        if (_tables.Count == 0 || source != _readFrom)
        {
            try
            {
                var loaded = await _host.RunAsync(
                    "Czytam model danych z przygotowanego skoroszytu",
                    (workspace, reporter) =>
                    {
                        WorkbookOutline outline = workspace.OpenForProcessing(config, reporter);
                        WorkbookInfo read = workspace.Inspect(config, reporter, readSplitValues: false);
                        return (Outline: outline, Info: read);
                    });

                _host.State.Outline = loaded.Outline;
                _tables = loaded.Info.ModelTables;
                _readFrom = source;
                _loaded = ("", "");
                _host.State.Keys = [];
            }
            catch (Exception ex)
            {
                _tables = [];
                _readFrom = ("", default);
                _loaded = ("", "");
                _host.State.Keys = [];

                _filling = true;
                try
                {
                    ListaTabel.ItemsSource = null;
                    ListaKolumn.ItemsSource = null;
                }
                finally
                {
                    _filling = false;
                }
                ShowValues();

                _host.Log(new ProgressLine(ProgressKind.Warning, $"Nie udało się odczytać modelu: {ex.Message}"));
                return;
            }
        }

        _filling = true;
        try
        {
            ListaTabel.ItemsSource = _tables.Select(table => table.Name).ToList();
            ListaTabel.SelectedItem = _tables
                .Select(table => table.Name)
                .FirstOrDefault(name => string.Equals(name, config.PowerPivot.SplitTable, StringComparison.OrdinalIgnoreCase));

            FillColumns(config.SplitColumn);
        }
        finally
        {
            _filling = false;
        }

        ShowValues();
        await EnsureValuesLoadedAsync();
    }

    public void Flush()
    {
        if (ListaTabel.SelectedItem is string table) _host.State.Config.PowerPivot.SplitTable = table;
        if (ListaKolumn.SelectedItem is string column) _host.State.Config.SplitColumn = column;
    }

    public async Task<string?> LeaveAsync()
    {
        if (ListaTabel.SelectedItem is not string) return "Wskaż tabelę modelu, w której jest kolumna podziału.";
        if (ListaKolumn.SelectedItem is not string) return "Wskaż kolumnę, po której dzielimy dane.";

        Flush();

        string? problem = await EnsureValuesLoadedAsync();
        if (problem is not null) return problem;

        if (_host.State.Keys.Count == 0)
            return $"W kolumnie '{_host.State.Config.SplitColumn}' nie ma żadnych wartości - nie da się po niej dzielić.";

        return null;
    }

    private void FillColumns(string? preferred)
    {
        List<string> columns = _tables
            .FirstOrDefault(table => string.Equals(table.Name, ListaTabel.SelectedItem as string, StringComparison.OrdinalIgnoreCase))
            ?.Columns.Select(column => column.Name).ToList() ?? [];

        ListaKolumn.ItemsSource = columns;
        ListaKolumn.SelectedItem = columns.FirstOrDefault(
            name => string.Equals(name, preferred, StringComparison.OrdinalIgnoreCase));
    }

    private void Tabela_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;

        _filling = true;
        try
        {
            FillColumns(null);
        }
        finally
        {
            _filling = false;
        }

        ShowValues();
    }

    private async void Kolumna_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;

        string? problem = await EnsureValuesLoadedAsync();
        if (problem is not null) _host.Log(new ProgressLine(ProgressKind.Warning, problem));
    }

    private async Task<string?> EnsureValuesLoadedAsync()
    {
        if (ListaTabel.SelectedItem is not string table || ListaKolumn.SelectedItem is not string column) return null;
        if (_loaded == (table, column)) return null;

        bool firstLoad = _loaded.Column.Length == 0;

        try
        {
            List<SplitKey> values = await _host.RunAsync(
                $"Czytam wartości {table}[{column}]",
                (workspace, reporter) => workspace.ReadModelValues(table, column, reporter));

            if (!firstLoad)
            {
                _host.State.Config.Values = [];
                _host.State.Config.Keys = [];
                _host.State.Config.Groups = [];
            }

            _host.State.Keys = SplitKeyCount.Unknown(values);
            _loaded = (table, column);
            ShowValues();
            return null;
        }
        catch (Exception ex)
        {
            _host.State.Keys = [];
            _loaded = ("", "");
            ShowValues();
            return $"Nie udało się odczytać {table}[{column}]: {ex.Message}";
        }
    }

    private static (string Path, DateTime WrittenUtc) SourceOf(SplitConfig config)
    {
        try
        {
            string path = config.WorkbookToProcess();
            return (path, File.GetLastWriteTimeUtc(path));
        }
        catch (Exception exception) when (exception is ArgumentException or IOException or NotSupportedException
                                              or UnauthorizedAccessException)
        {
            return ("", default);
        }
    }

    private void ShowValues()
    {
        List<SplitKeyCount> keys = _host.State.Keys;
        ListaWartosci.ItemsSource = keys;
        Podsumowanie.Text = keys.Count == 0
            ? ""
            : $"{keys.Count} różnych wartości - tyle plików powstanie, jeśli nie złożysz ich w grupy.";
    }
}
