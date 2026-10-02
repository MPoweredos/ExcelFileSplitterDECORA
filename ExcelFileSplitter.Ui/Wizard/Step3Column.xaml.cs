using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step3Column : UserControl, IWizardStep
{
    private const string Zaden = "— nie używam —";

    private readonly IWizardHost _host;

    private List<string> _loadedColumns = [];

    private bool _filling;

    public Step3Column(IWizardHost host)
    {
        _host = host;
        InitializeComponent();

        PasekPodzialow.Attach(host);
        PasekPodzialow.Saving += Flush;
        PasekPodzialow.Changed += PassChanged;
    }

    private async void PassChanged()
    {
        ShowCurrentPass();

        string? problem = await EnsureKeysLoadedAsync();
        if (problem is not null) _host.Log(new ProgressLine(ProgressKind.Warning, problem));
    }

    public string Title => "3. Kolumny podziału";

    public async Task EnterAsync()
    {
        ShowCurrentPass();

        string? problem = await EnsureKeysLoadedAsync();
        if (problem is not null) _host.Log(new ProgressLine(ProgressKind.Warning, problem));
    }

    private void ShowCurrentPass()
    {
        List<string> headers = _host.State.Data?.Headers.Where(header => header.Length > 0).ToList() ?? [];
        IReadOnlyList<string> configured = _host.State.CurrentPass.SplitColumnNames;

        PasekPodzialow.Refresh();

        _filling = true;
        try
        {
            ListaKolumn1.ItemsSource = headers;
            ListaKolumn2.ItemsSource = WithNone(headers);
            ListaKolumn3.ItemsSource = WithNone(headers);

            Select(ListaKolumn1, configured.Count > 0 ? configured[0] : null, headers);
            Select(ListaKolumn2, configured.Count > 1 ? configured[1] : null, headers);
            Select(ListaKolumn3, configured.Count > 2 ? configured[2] : null, headers);
        }
        finally
        {
            _filling = false;
        }

        _loadedColumns = _host.State.Keys.Count > 0 ? configured.ToList() : [];

        ShowKeys();
    }

    public void Flush()
    {
        SplitPass pass = _host.State.CurrentPass;
        List<string> columns = ChosenColumns();
        if (columns.Count > 0)
        {
            if (columns.Count == 1)
            {
                pass.SplitColumn = columns[0];
                pass.SplitColumns = [];
            }
            else
            {
                pass.SplitColumn = "";
                pass.SplitColumns = columns;
            }
        }

        _host.State.ApplyToConfig();
    }

    public async Task<string?> LeaveAsync()
    {
        List<string> columns = ChosenColumns();

        if (columns.Count == 0) return "Wskaż przynajmniej jedną kolumnę, po której dzielimy dane.";

        string? duplicate = columns
            .GroupBy(column => column, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1)?.Key;

        if (duplicate is not null)
            return $"Kolumna '{duplicate}' została wybrana więcej niż raz - każda kolumna podziału może być tylko jedna.";

        Flush();

        if (_host.State.ManyPasses)
        {
            List<SplitPass> passes = _host.State.Drafts.Select(draft => draft.Pass).ToList();

            if (passes.Any(pass => pass.Name.Trim().Length == 0))
                return "Przy kilku podziałach każdy musi mieć nazwę - trafia ona do nazwy pliku i mówi, " +
                       "z którego podziału plik pochodzi.";

            string? duplicatePass = passes
                .GroupBy(pass => pass.Name.Trim(), StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => group.Count() > 1)?.Key;

            if (duplicatePass is not null)
                return $"Podział o nazwie '{duplicatePass}' jest więcej niż raz - nazwy muszą się różnić.";
        }

        string? problem = await EnsureKeysLoadedAsync();
        if (problem is not null) return problem;

        if (_host.State.Keys.Count == 0)
            return $"W {(columns.Count == 1 ? $"kolumnie '{columns[0]}'" : "wybranych kolumnach")} nie ma żadnych " +
                   "kompletnych wartości - nie da się po nich dzielić.";

        return null;
    }

    private static List<string> WithNone(IEnumerable<string> headers) => [Zaden, .. headers];

    private static void Select(ComboBox list, string? column, IReadOnlyList<string> headers)
    {
        if (column is null || !headers.Contains(column, StringComparer.OrdinalIgnoreCase))
        {
            list.SelectedIndex = list.ItemsSource is List<string> { Count: > 0 } items && items[0] == Zaden ? 0 : -1;
            return;
        }

        list.SelectedItem = ((List<string>)list.ItemsSource)
            .First(item => string.Equals(item, column, StringComparison.OrdinalIgnoreCase));
    }

    private List<string> ChosenColumns()
    {
        var columns = new List<string>();

        if (ListaKolumn1.SelectedItem is string first && first.Length > 0) columns.Add(first);
        if (ListaKolumn2.SelectedIndex > 0 && ListaKolumn2.SelectedItem is string second) columns.Add(second);
        if (ListaKolumn3.SelectedIndex > 0 && ListaKolumn3.SelectedItem is string third) columns.Add(third);

        return columns;
    }

    private async void Kolumna_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;

        string? problem = await EnsureKeysLoadedAsync();
        if (problem is not null) _host.Log(new ProgressLine(ProgressKind.Warning, problem));
    }

    private async Task<string?> EnsureKeysLoadedAsync()
    {
        List<string> columns = ChosenColumns();
        if (columns.Count == 0) return null;
        if (_loadedColumns.SequenceEqual(columns, StringComparer.OrdinalIgnoreCase)) return null;

        WorksheetOptions source = _host.State.Config.Worksheet;
        bool firstLoad = _loadedColumns.Count == 0;

        try
        {
            _host.State.Keys = await _host.RunAsync(
                $"Czytam kombinacje: {string.Join(" + ", columns)}",
                (workspace, _) => workspace.ReadSplitKeys(source, columns));

            if (!firstLoad)
            {
                SplitPass pass = _host.State.CurrentPass;
                pass.Values = [];
                pass.Keys = [];
                pass.Groups = [];
            }

            _loadedColumns = columns;
            ShowKeys();
            return null;
        }
        catch (Exception ex)
        {
            _host.State.Keys = [];
            _loadedColumns = [];
            ShowKeys();
            return $"Nie udało się odczytać kolumn ({string.Join(" + ", columns)}): {ex.Message}";
        }
    }

    private void ShowKeys()
    {
        List<SplitKeyCount> keys = _host.State.Keys;
        List<string> columns = ChosenColumns();

        NaglowekListy.Text = columns.Count > 1 ? "Kombinacje w danych:" : "Wartości w danych:";
        ListaWartosci.ItemsSource = keys;

        if (keys.Count == 0)
        {
            Podsumowanie.Text = "";
            return;
        }

        long? rows = keys.Sum(key => key.RowCount);
        Podsumowanie.Text = $"{keys.Count} {(columns.Count > 1 ? "kombinacji" : "wartości")}" +
                            (rows is null ? "" : $", razem {rows:N0} wierszy") +
                            " - tyle plików powstanie, jeśli nie złożysz ich w grupy.";
    }
}
