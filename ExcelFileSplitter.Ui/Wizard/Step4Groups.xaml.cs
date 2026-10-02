using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step4Groups : UserControl, IWizardStep
{
    private readonly IWizardHost _host;
    private readonly List<RecipientGroup> _groups = [];

    private RecipientGroup? _editing;

    private bool _filling;

    private readonly HashSet<RecipientGroup> _namedByUser = [];

    public Step4Groups(IWizardHost host)
    {
        _host = host;
        InitializeComponent();

        PasekPodzialow.Attach(host);
        PasekPodzialow.Saving += Flush;
        PasekPodzialow.Changed += () => _ = EnterAsync();

        PasekPodzialow.Added += () => _host.GoToStep<Step3Column>();
    }

    public string Title => "4. Grupy i arkusze";

    public Task EnterAsync()
    {
        _filling = true;
        try
        {
            SplitPass pass = _host.State.CurrentPass;

            _groups.Clear();
            _namedByUser.Clear();
            foreach (RecipientGroup group in pass.Groups)
            {
                RecipientGroup copy = Copy(group);
                _groups.Add(copy);
                if (copy.Name.Length > 0) _namedByUser.Add(copy);
            }
            ListaGrup.ItemsSource = null;
            ListaGrup.ItemsSource = _groups;

            ListaWartosci.ItemsSource = KeyItems(pass.RequestedKeys());
            NaglowekWartosci.Text = pass.SplitColumnNames.Count > 1
                ? "Dla których kombinacji utworzyć pliki (nic nie zaznaczone = dla wszystkich):"
                : "Dla których wartości utworzyć pliki (nic nie zaznaczone = dla wszystkich):";
            ListaArkuszyWspolnych.ItemsSource = SheetItems(pass.SheetsToKeep);

            bool groupMode = _groups.Count > 0;
            TrybGrup.IsChecked = groupMode;
            TrybNaWartosc.IsChecked = !groupMode;
        }
        finally
        {
            _filling = false;
        }

        ShowMode();
        PasekPodzialow.Refresh();
        SelectGroup(_groups.FirstOrDefault());
        return Task.CompletedTask;
    }

    public void Flush()
    {
        SaveEditor();

        SplitPass pass = _host.State.CurrentPass;
        pass.SheetsToKeep = Chosen(ListaArkuszyWspolnych);

        if (TrybGrup.IsChecked == true)
        {
            pass.Groups = _groups.Select(Copy).ToList();
            pass.Values = [];
            pass.Keys = [];
        }
        else
        {
            pass.Groups = [];
            WriteKeys(ChosenKeys(ListaWartosci), values => pass.Values = values, keys => pass.Keys = keys);
        }

        _host.State.ApplyToConfig();
    }

    public Task<string?> LeaveAsync()
    {
        Flush();
        SplitPass pass = _host.State.CurrentPass;

        if (TrybGrup.IsChecked == true)
        {
            if (_groups.Count == 0)
                return Refuse("Dodaj przynajmniej jeden plik zbiorczy albo wróć do podziału na każdą wartość.");

            RecipientGroup? noName = _groups.FirstOrDefault(group => group.Name.Trim().Length == 0);
            if (noName is not null) return Refuse("Każdy plik zbiorczy musi mieć nazwę - trafia ona do nazwy pliku.");

            RecipientGroup? empty = _groups.FirstOrDefault(group => group.RequestedKeys().Count == 0);
            if (empty is not null)
                return Refuse($"Plik '{empty.Name}' nie ma przypisanej żadnej wartości - powstałby pusty.");
        }

        if (pass.SheetsToKeep.Count == 0 && !pass.Groups.Any(group => group.SheetsToKeep.Count > 0))
            return Refuse("Zaznacz arkusze, które mają zostać w plikach - inaczej odbiorcy dostaną puste skoroszyty.");

        return Task.FromResult<string?>(null);
    }

    private static Task<string?> Refuse(string reason) => Task.FromResult<string?>(reason);

    private static RecipientGroup Copy(RecipientGroup group) => new()
    {
        Name = group.Name,
        Values = [.. group.Values],
        Keys = group.Keys.Select(key => key.ToList()).ToList(),
        SheetsToKeep = [.. group.SheetsToKeep],
    };

    private List<PickItem> KeyItems(IEnumerable<SplitKey> selected)
    {
        var chosen = selected.ToHashSet();
        return _host.State.Keys
            .Select(key => new PickItem
            {
                Text = key.Display,
                Note = key.RowCountLabel,
                IsSelected = chosen.Contains(key.Key),
                Tag = key.Key,
            })
            .ToList();
    }

    private static List<SplitKey> ChosenKeys(ItemsControl list) =>
        (list.ItemsSource as IEnumerable<PickItem> ?? [])
        .Where(item => item.IsSelected)
        .Select(item => item.Tag)
        .OfType<SplitKey>()
        .ToList();

    private static void WriteKeys(List<SplitKey> keys, Action<List<string>> setValues, Action<List<List<string>>> setKeys)
    {
        bool single = keys.All(key => key.Count == 1);
        setValues(single ? keys.Select(key => key.Values[0]).ToList() : []);
        setKeys(single ? [] : keys.Select(key => key.Values.ToList()).ToList());
    }

    private List<PickItem> SheetItems(IEnumerable<string> selected)
    {
        var chosen = selected.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var hidden = (_host.State.Outline?.Sheets ?? [])
            .Where(sheet => !sheet.IsVisible)
            .Select(sheet => sheet.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        string? dataSheet = _host.State.Data?.SheetName;

        return _host.State.SheetNames
            .Select(name => new PickItem
            {
                Text = name,
                Note = Note(name),
                IsSelected = chosen.Contains(name),
            })
            .ToList();

        string Note(string name)
        {
            var notes = new List<string>();
            if (hidden.Contains(name)) notes.Add("ukryty");
            if (string.Equals(name, dataSheet, StringComparison.OrdinalIgnoreCase))
                notes.Add("arkusz z danymi - można go usunąć, tabele przestawne mają własny cache");
            return notes.Count == 0 ? "" : "(" + string.Join(", ", notes) + ")";
        }
    }

    private static List<string> Chosen(ItemsControl list) =>
        (list.ItemsSource as IEnumerable<PickItem> ?? [])
        .Where(item => item.IsSelected)
        .Select(item => item.Text)
        .ToList();

    private void ShowMode()
    {
        bool groupMode = TrybGrup.IsChecked == true;
        PanelGrup.Visibility = groupMode ? Visibility.Visible : Visibility.Collapsed;
        PanelWartosci.Visibility = groupMode ? Visibility.Collapsed : Visibility.Visible;
    }

    private void Tryb_Changed(object sender, RoutedEventArgs e)
    {
        if (_filling) return;
        ShowMode();
    }

    private void Grupa_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        SaveEditor();
        SelectGroup(ListaGrup.SelectedItem as RecipientGroup);
    }

    private void SaveEditor()
    {
        if (_editing is null) return;

        RecipientGroup group = _editing;
        group.Name = PoleNazwaGrupy.Text.Trim();
        group.SheetsToKeep = Chosen(ListaArkuszyGrupy);
        WriteKeys(ChosenKeys(ListaWartosciGrupy), values => group.Values = values, keys => group.Keys = keys);
    }

    private static string AutoName(IReadOnlyList<SplitKey> keys) => keys.Count switch
    {
        0 => "",
        1 => keys[0].Display,
        2 => $"{keys[0].Display} + {keys[1].Display}",
        _ => $"{keys[0].Display} + {keys[1].Display} i {keys.Count - 2} inn.",
    };

    private void RefreshAutoName()
    {
        if (_editing is null || _namedByUser.Contains(_editing)) return;

        string name = AutoName(ChosenKeys(ListaWartosciGrupy));
        if (name.Length == 0) return;

        _editing.Name = name;

        _filling = true;
        try
        {
            PoleNazwaGrupy.Text = name;
            RefreshGroupList();
        }
        finally
        {
            _filling = false;
        }
    }

    private void RefreshFileNamePreview()
    {
        string name = PoleNazwaGrupy.Text.Trim();

        PodgladNazwyPliku.Text = name.Length == 0
            ? "Wpisz nazwę - to ona trafia do nazwy pliku."
            : $"Powstanie plik: {_host.State.Config.ForPass(_host.State.CurrentPass).OutputFileNameFor(name)}";
    }

    private void RefreshGroupList()
    {
        object? selected = ListaGrup.SelectedItem;
        ListaGrup.ItemsSource = null;
        ListaGrup.ItemsSource = _groups;
        ListaGrup.SelectedItem = selected;
    }

    private void SelectGroup(RecipientGroup? group)
    {
        _editing = group;
        _filling = true;
        try
        {
            ListaGrup.SelectedItem = group;
            EdytorGrupy.Visibility = group is null ? Visibility.Hidden : Visibility.Visible;
            PrzyciskUsun.IsEnabled = group is not null;

            PoleNazwaGrupy.Text = group?.Name ?? "";
            ListaWartosciGrupy.ItemsSource = KeyItems(group?.RequestedKeys() ?? []);
            ListaArkuszyGrupy.ItemsSource = SheetItems(group?.SheetsToKeep ?? []);
            RefreshFileNamePreview();
        }
        finally
        {
            _filling = false;
        }
    }

    private void NazwaGrupy_Changed(object sender, TextChangedEventArgs e)
    {
        if (_filling || _editing is null) return;

        _namedByUser.Add(_editing);
        _editing.Name = PoleNazwaGrupy.Text.Trim();

        _filling = true;
        try
        {
            RefreshGroupList();
        }
        finally
        {
            _filling = false;
        }

        RefreshFileNamePreview();
    }

    private void Ptaszek_Click(object sender, RoutedEventArgs e)
    {
        if (_filling || _editing is null || TrybGrup.IsChecked != true) return;

        RefreshAutoName();
        RefreshFileNamePreview();
    }

    private void NazwaZWartosci_Click(object sender, RoutedEventArgs e)
    {
        if (_editing is null) return;

        List<SplitKey> keys = ChosenKeys(ListaWartosciGrupy);
        if (keys.Count == 0)
        {
            PodgladNazwyPliku.Text = "Najpierw zaznacz po lewej, czyje dane trafią do tego pliku.";
            return;
        }

        _namedByUser.Remove(_editing);
        RefreshAutoName();
        RefreshFileNamePreview();
    }

    private void Dodaj_Click(object sender, RoutedEventArgs e)
    {
        SaveEditor();

        var group = new RecipientGroup { Name = $"Plik {_groups.Count + 1}" };
        _groups.Add(group);

        _filling = true;
        try
        {
            ListaGrup.ItemsSource = null;
            ListaGrup.ItemsSource = _groups;
        }
        finally
        {
            _filling = false;
        }

        SelectGroup(group);
    }

    private void Usun_Click(object sender, RoutedEventArgs e)
    {
        if (ListaGrup.SelectedItem is not RecipientGroup group) return;

        _groups.Remove(group);
        _namedByUser.Remove(group);
        _editing = null;

        _filling = true;
        try
        {
            ListaGrup.ItemsSource = null;
            ListaGrup.ItemsSource = _groups;
        }
        finally
        {
            _filling = false;
        }

        SelectGroup(_groups.FirstOrDefault());
    }
}
