using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using Microsoft.Win32;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step1File : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    public Step1File(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "1. Plik do podziału";

    public Task EnterAsync()
    {
        SplitConfig config = _host.State.Config;
        PoleZrodlo.Text = config.SourceWorkbook;
        PoleWynik.Text = config.OutputFolder;
        PoleSzablon.Text = config.FileNameTemplate;
        return Task.CompletedTask;
    }

    public void Flush()
    {
        SplitConfig config = _host.State.Config;
        config.SourceWorkbook = PoleZrodlo.Text.Trim();
        config.OutputFolder = PoleWynik.Text.Trim();
        config.FileNameTemplate = PoleSzablon.Text.Trim();
    }

    public async Task<string?> LeaveAsync()
    {
        Flush();
        SplitConfig config = _host.State.Config;

        if (config.SourceWorkbook.Length == 0) return "Wskaż plik do podziału.";
        if (!File.Exists(config.SourceWorkbook)) return $"Nie ma pliku {config.SourceWorkbook}.";

        if (!ExcelFileTypes.IsWorkbook(config.SourceWorkbook))
            return $"{Path.GetFileName(config.SourceWorkbook)} to nie jest plik Excela. " +
                   $"Wskaż plik {ExcelFileTypes.Listed}.";
        if (config.OutputFolder.Length == 0) return "Wskaż folder, w którym mają powstać pliki wynikowe.";
        if (!config.FileNameTemplate.Contains("{value}") && !config.FileNameTemplate.Contains("{name}"))
            return "Nazwa pliku musi zawierać {value} albo {name} - inaczej każdy odbiorca nadpisałby poprzedniego.";

        if (_host.State.OutlineMatchesSource) return null;

        WorkbookOutline outline;
        try
        {
            outline = await _host.RunAsync(
                $"Otwieram {Path.GetFileName(config.SourceWorkbook)} - przy dużym pliku potrwa to kilkanaście sekund",
                (workspace, reporter) => workspace.OpenWorkbook(config.SourceWorkbook, reporter));
        }
        catch (Exception ex)
        {
            return $"Nie udało się otworzyć pliku: {ex.Message}";
        }

        if (outline.Candidates.Count == 0)
            return "W tym skoroszycie nie ma ani tabeli Excela, ani arkusza z nagłówkami - nie ma czego dzielić.";

        _host.State.Outline = outline;
        _host.State.Data = null;
        _host.State.ForgetKeys();
        return null;
    }

    private void Zrodlo_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Plik do podziału",
            Filter = "Pliki Excela (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|Wszystkie pliki|*.*",
            CheckFileExists = true,
        };
        if (PoleZrodlo.Text.Length > 0) TrySetInitialDirectory(dialog, PoleZrodlo.Text);
        if (dialog.ShowDialog() != true) return;

        PoleZrodlo.Text = dialog.FileName;
        if (PoleWynik.Text.Length == 0)
            PoleWynik.Text = Path.Combine(Path.GetDirectoryName(dialog.FileName) ?? "", "Podzielone");
    }

    private void Wynik_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "Folder na pliki wynikowe" };
        if (PoleWynik.Text.Length > 0) dialog.InitialDirectory = PoleWynik.Text;
        if (dialog.ShowDialog() == true) PoleWynik.Text = dialog.FolderName;
    }

    private static void TrySetInitialDirectory(OpenFileDialog dialog, string path)
    {
        try
        {
            string? directory = Path.GetDirectoryName(path);
            if (directory is not null && Directory.Exists(directory)) dialog.InitialDirectory = directory;
        }
        catch (Exception exception) when (exception is ArgumentException or PathTooLongException)
        {
        }
    }
}
