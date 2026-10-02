using System.IO;
using System.Windows;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class LoadedConfigDialog : Window
{
    public LoadedConfigDialog(SplitConfig config, string configPath)
    {
        InitializeComponent();

        string source = config.SourceWorkbook.Length > 0 ? config.SourceWorkbook : "(nie wskazano pliku źródłowego)";
        OpisPliku.Text = $"{Path.GetFileName(configPath)}\n" +
                         $"Plik do podziału: {source}\n" +
                         $"Podział według: {(config.SplitColumnNames.Count > 0 ? string.Join(" + ", config.SplitColumnNames) : "(nie wskazano)")}";

        OstrzezenieKontroli.Visibility = config.Worksheet.DeleteEntireRow ? Visibility.Visible : Visibility.Collapsed;
    }

    public bool SkipFileReview => PomijajPrzegladPliku.IsChecked == true;

    public bool SkipReadinessCheck => PomijajKontrole.IsChecked == true;

    private void Dalej_Click(object sender, RoutedEventArgs e) => DialogResult = true;

    private void Anuluj_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
