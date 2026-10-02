using System.Collections;
using System.Collections.Specialized;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class LogView : UserControl
{
    private INotifyCollectionChanged? _watched;

    public LogView()
    {
        InitializeComponent();
        Loaded += (_, _) => Watch();
        Unloaded += (_, _) => Unwatch();
    }

    public IEnumerable? ItemsSource
    {
        get => Lista.ItemsSource;
        set
        {
            Unwatch();
            Lista.ItemsSource = value;
            if (IsLoaded) Watch();
        }
    }

    private void Watch()
    {
        Unwatch();
        if (Lista.ItemsSource is not INotifyCollectionChanged source) return;

        source.CollectionChanged += Lines_Changed;
        _watched = source;
    }

    private void Unwatch()
    {
        if (_watched is null) return;

        _watched.CollectionChanged -= Lines_Changed;
        _watched = null;
    }

    private void Lines_Changed(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (!IsLoaded) return;
        if (e.Action != NotifyCollectionChangedAction.Add) return;

        if (e.NewItems is null || e.NewItems.Count == 0) return;

        object? last = e.NewItems[e.NewItems.Count - 1];
        if (last is not null) Lista.ScrollIntoView(last);
    }

    private void Lista_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.C || (Keyboard.Modifiers & ModifierKeys.Control) == 0) return;

        CopyToClipboard(SelectedOrAll());
        e.Handled = true;
    }

    private void KopiujZaznaczone_Click(object sender, RoutedEventArgs e) => CopyToClipboard(SelectedOrAll());

    private void KopiujWszystko_Click(object sender, RoutedEventArgs e) => CopyToClipboard(All());

    private void ZaznaczWszystko_Click(object sender, RoutedEventArgs e) => Lista.SelectAll();

    private List<ProgressLine> SelectedOrAll()
    {
        var selected = Lista.SelectedItems.OfType<ProgressLine>().ToList();
        return selected.Count > 0 ? selected : All();
    }

    private List<ProgressLine> All() => (Lista.ItemsSource ?? Array.Empty<object>()).OfType<ProgressLine>().ToList();

    private void CopyToClipboard(List<ProgressLine> lines)
    {
        if (lines.Count == 0)
        {
            Potwierdzenie.Text = "Nie ma czego kopiować.";
            return;
        }

        var text = new StringBuilder();
        foreach (ProgressLine line in lines) text.AppendLine(line.Display);

        try
        {
            Clipboard.SetText(text.ToString());
            Potwierdzenie.Text = $"Skopiowano {lines.Count} lini(i).";
        }
        catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException or InvalidOperationException)
        {
            Potwierdzenie.Text = "Nie udało się otworzyć schowka - spróbuj jeszcze raz.";
        }
    }
}
