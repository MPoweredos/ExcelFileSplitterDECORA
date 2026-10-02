using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using Microsoft.Win32;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PivotStep2Queries : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    public PivotStep2Queries(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "2. Zapytania i parametry";

    public Task EnterAsync()
    {
        SplitConfig config = _host.State.Config;

        ListaZapytan.ItemsSource = config.PowerPivot.Queries.Select(query => new QueryRow(
            query.Name,
            string.IsNullOrWhiteSpace(query.FilterColumn) ? "(bez filtra)" : query.FilterColumn,
            query.KeyFrom is null ? "" : $"{query.KeyFrom.Query}[{query.KeyFrom.Column}]",
            query.Cache ? "tak" : "nie")).ToList();

        bool empty = config.PowerPivot.Queries.Count == 0;
        RamkaBraku.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
        NaglowekListy.Text = empty
            ? "Zapytania z wczytanych ustawień: (brak)"
            : $"Zapytania z wczytanych ustawień ({config.PowerPivot.Queries.Count}):";

        OpisParametrow.Text = $"Tabela parametrów: {config.PowerPivot.ParamsTable} " +
                              $"({config.PowerPivot.SplitParamName}, {config.PowerPivot.UseCacheParamName}, {config.PowerPivot.CachePathParamName}), " +
                              $"arkusz parametrów: {config.PowerPivot.ConfigSheetName}";

        return Task.CompletedTask;
    }

    public void Flush() { }

    public Task<string?> LeaveAsync()
    {
        if (_host.State.Config.PowerPivot.Queries.Count == 0)
            return Task.FromResult<string?>(
                "Bez zapytań przygotowanie nie ma czego filtrować - każdy plik dostałby komplet danych. " +
                "Zapisz szablon, wypełnij sekcję Queries i wczytaj ustawienia przyciskiem na dole okna.");

        return Task.FromResult<string?>(null);
    }

    private void Szablon_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Zapisz szablon ustawień",
            Filter = "Ustawienia podziału (*.json)|*.json",
            FileName = "config.json",
            AddExtension = true,
            DefaultExt = ".json",
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            SplitConfig template = SplitConfig.CreateSample();

            SplitConfig current = _host.State.Config;
            if (current.SourceWorkbook.Length > 0) template.SourceWorkbook = current.SourceWorkbook;
            if (current.OutputFolder.Length > 0) template.OutputFolder = current.OutputFolder;
            if (current.PowerPivot.CacheFolder.Length > 0) template.PowerPivot.CacheFolder = current.PowerPivot.CacheFolder;

            template.SaveTemplate(dialog.FileName);
            _host.Log(new Engine.ProgressLine(Engine.ProgressKind.Step, $"Zapisano szablon: {dialog.FileName}"));
            OpisParametrow.Text = $"Zapisano szablon do {Path.GetFileName(dialog.FileName)}. " +
                                  "Wypełnij w nim Queries i wczytaj go przyciskiem \"Wczytaj ustawienia\" na dole okna.";
        }
        catch (Exception ex)
        {
            _host.Log(new Engine.ProgressLine(Engine.ProgressKind.Warning, $"Nie udało się zapisać szablonu: {ex.Message}"));
        }
    }

    public sealed record QueryRow(string Name, string Filter, string Key, string Cache);
}
