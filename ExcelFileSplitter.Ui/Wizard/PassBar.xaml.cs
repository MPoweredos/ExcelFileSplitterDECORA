using System.Windows;
using System.Windows.Controls;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class PassBar : UserControl
{
    private IWizardHost? _host;

    private bool _filling;

    public PassBar()
    {
        InitializeComponent();
    }

    public event Action? Saving;

    public event Action? Changed;

    public event Action? Added;

    public void Attach(IWizardHost host)
    {
        _host = host;
        Refresh();
    }

    public void Refresh()
    {
        if (_host is null) return;

        WizardState state = _host.State;

        if (state.Config.Mode != SplitMode.Worksheet)
        {
            Ramka.Visibility = Visibility.Collapsed;
            return;
        }

        Ramka.Visibility = Visibility.Visible;

        _filling = true;
        try
        {
            Lista.ItemsSource = null;
            Lista.ItemsSource = state.Drafts;
            Lista.SelectedIndex = state.Current;
            PoleNazwa.Text = state.CurrentPass.Name;
        }
        finally
        {
            _filling = false;
        }

        int count = state.Drafts.Count;
        Lista.IsEnabled = count > 1;
        PrzyciskUsun.IsEnabled = count > 1;
        Etykieta.Text = count > 1 ? $"Podział ({state.Current + 1} z {count})" : "Podział";

        Opis.Text = count == 1
            ? "\"Dodaj podział\" tworzy DRUGI, niezależny podział tego samego pliku - np. pierwszy po menadżerze, " +
              "drugi po obsłudze klienta. Pliki z obu powstają w jednym przebiegu i się SUMUJĄ, a nie mnożą."
            : "Każdy podział ma własne kolumny, wartości, grupy i arkusze. Przełączaj się między nimi listą obok - " +
              "oba kroki kreatora dotyczą tego, który jest tu wybrany.";
    }

    private void Lista_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filling || _host is null) return;

        int chosen = Lista.SelectedIndex;
        if (chosen < 0 || chosen == _host.State.Current) return;

        Saving?.Invoke();
        _host.State.Current = chosen;
        Refresh();
        Changed?.Invoke();
    }

    private void Nazwa_Changed(object sender, TextChangedEventArgs e)
    {
        if (_filling || _host is null) return;

        _host.State.CurrentPass.Name = PoleNazwa.Text.Trim();

        int selected = Lista.SelectedIndex;
        _filling = true;
        try
        {
            Lista.ItemsSource = null;
            Lista.ItemsSource = _host.State.Drafts;
            Lista.SelectedIndex = selected;
        }
        finally
        {
            _filling = false;
        }
    }

    private void Dodaj_Click(object sender, RoutedEventArgs e)
    {
        if (_host is null) return;

        Saving?.Invoke();
        _host.State.AddPass();
        Refresh();
        Changed?.Invoke();
        Added?.Invoke();
    }

    private void Usun_Click(object sender, RoutedEventArgs e)
    {
        if (_host is null || !_host.State.RemovePass(_host.State.Current)) return;

        _host.State.ApplyToConfig();
        Refresh();
        Changed?.Invoke();
    }
}
