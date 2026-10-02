using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;
using Microsoft.Win32;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PivotStep1Files : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    public PivotStep1Files(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "1. Plik i foldery";

    public Task EnterAsync()
    {
        SplitConfig config = _host.State.Config;
        PoleZrodlo.Text = config.SourceWorkbook;
        PoleWynik.Text = config.OutputFolder;
        PoleCache.Text = config.PowerPivot.CacheFolder;
        PoleSzablon.Text = config.FileNameTemplate;

        ShowPreparedPath();
        return Task.CompletedTask;
    }

    public void Flush()
    {
        SplitConfig config = _host.State.Config;
        config.SourceWorkbook = PoleZrodlo.Text.Trim();
        config.OutputFolder = PoleWynik.Text.Trim();
        config.PowerPivot.CacheFolder = PoleCache.Text.Trim();
        config.FileNameTemplate = PoleSzablon.Text.Trim();
    }

    public Task<string?> LeaveAsync()
    {
        Flush();
        SplitConfig config = _host.State.Config;

        if (config.SourceWorkbook.Length == 0) return Refuse("Wskaż plik do podziału.");
        if (!File.Exists(config.SourceWorkbook)) return Refuse($"Nie ma pliku {config.SourceWorkbook}.");
        if (!ExcelFileTypes.CanHoldDataModel(config.SourceWorkbook))
            return Refuse($"{Path.GetFileName(config.SourceWorkbook)} to nie jest plik, który może trzymać " +
                          "model danych. Wskaż plik xlsx, xlsm albo xlsb.");
        if (config.OutputFolder.Length == 0) return Refuse("Wskaż folder, w którym mają powstać pliki wynikowe.");
        if (!config.FileNameTemplate.Contains("{value}") && !config.FileNameTemplate.Contains("{name}"))
            return Refuse("Nazwa pliku musi zawierać {value} albo {name} - inaczej każdy odbiorca nadpisałby poprzedniego.");

        return Task.FromResult<string?>(null);
    }

    private static Task<string?> Refuse(string reason) => Task.FromResult<string?>(reason);

    private void Zrodlo_Changed(object sender, TextChangedEventArgs e) => ShowPreparedPath();

    private void ShowPreparedPath()
    {
        string source = PoleZrodlo.Text.Trim();
        if (source.Length == 0)
        {
            SciezkaRobocza.Text = "(najpierw wskaż plik źródłowy)";
            return;
        }

        try
        {
            SplitConfig config = _host.State.Config;
            string prepared = config.PowerPivot.PreparedWorkbook.Length > 0
                ? config.PowerPivot.PreparedWorkbook
                : Path.Combine(
                    Path.GetDirectoryName(Path.GetFullPath(source)) ?? "",
                    Path.GetFileNameWithoutExtension(source) + ".prepared" + Path.GetExtension(source));

            SciezkaRobocza.Text = prepared + (File.Exists(prepared) ? "   (już istnieje)" : "   (jeszcze nie istnieje)");
        }
        catch (Exception exception) when (exception is ArgumentException or PathTooLongException or NotSupportedException)
        {
            SciezkaRobocza.Text = "";
        }
    }

    private void Zrodlo_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Raport z Modelem danych",
            Filter = "Pliki z modelem danych (*.xlsm;*.xlsx;*.xlsb)|*.xlsm;*.xlsx;*.xlsb|Wszystkie pliki|*.*",
            CheckFileExists = true,
        };
        if (dialog.ShowDialog() != true) return;

        PoleZrodlo.Text = dialog.FileName;
        string? folder = Path.GetDirectoryName(dialog.FileName);
        if (folder is null) return;

        if (PoleWynik.Text.Length == 0) PoleWynik.Text = Path.Combine(folder, "Podzielone");
        if (PoleCache.Text.Length == 0) PoleCache.Text = Path.Combine(folder, "cache");
    }

    private void Wynik_Click(object sender, RoutedEventArgs e) => Pick(PoleWynik, "Folder na pliki wynikowe");

    private void Cache_Click(object sender, RoutedEventArgs e) => Pick(PoleCache, "Folder na cache");

    private static void Pick(TextBox target, string title)
    {
        var dialog = new OpenFolderDialog { Title = title };
        if (target.Text.Length > 0) dialog.InitialDirectory = target.Text;
        if (dialog.ShowDialog() == true) target.Text = dialog.FolderName;
    }
}
