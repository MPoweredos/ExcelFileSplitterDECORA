using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class RecipientsTests
{
    private static SplitKey Klucz(params string[] values) => SplitKey.Of(values);

    private static List<SplitKey> Klucze(params string[] values) => values.Select(value => Klucz(value)).ToList();

    private static SplitConfig Config() => new()
    {
        SourceWorkbook = @"C:\raporty\Raport.xlsm",
        SplitColumn = "Opiekun",
        FileNameTemplate = "Raport_{value}.xlsx",
        PowerPivot = new PowerPivotOptions { SplitTable = "dim_Klienci" },
    };

    [Fact]
    public void Bez_Values_i_Groups_powstaje_plik_na_kazda_wartosc_z_modelu()
    {
        RecipientPlan plan = Recipients.BuildPlan(Config(), Klucze("Jan Kowalski", "Anna Nowak"));

        Assert.Equal(2, plan.Files.Count);
        Assert.Empty(plan.Problems);
        Assert.Empty(plan.Warnings);
        Assert.Equal("Jan Kowalski", plan.Files[0].Name);
        Assert.Equal(new[] { "Jan Kowalski" }, plan.Files[0].Keys.Select(key => key.Display));
        Assert.False(plan.Files[0].IsGroup);
    }

    [Fact]
    public void Kolejnosc_plikow_idzie_za_kolejnoscia_wartosci_z_modelu()
    {
        RecipientPlan plan = Recipients.BuildPlan(Config(), Klucze("Anna Nowak", "Jan Kowalski"));

        Assert.Equal("Anna Nowak", plan.Files[0].Name);
        Assert.Equal("Jan Kowalski", plan.Files[1].Name);
    }

    [Fact]
    public void Wartosc_z_konfiguracji_dostaje_pisownie_z_modelu()
    {
        var config = Config();
        config.Values = ["jan KOWALSKI"];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski", "Anna Nowak"));

        Assert.Single(plan.Files);
        Assert.Equal(new[] { "Jan Kowalski" }, plan.Files[0].Keys.Select(key => key.Display));
        Assert.Empty(plan.Problems);
    }

    [Fact]
    public void Wartosc_spoza_modelu_blokuje_start()
    {
        var config = Config();
        config.Values = ["Duch"];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Single(plan.Problems);
        Assert.Contains("Duch", plan.Problems[0]);
    }

    [Fact]
    public void Grupa_z_czesciowo_nieznanymi_wartosciami_daje_ostrzezenie_a_nie_blad()
    {
        var config = Config();
        config.Groups = [new RecipientGroup { Name = "Menadzer", Values = ["Jan Kowalski", "Duch"] }];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Empty(plan.Problems);
        Assert.Single(plan.Warnings);
        Assert.Contains("Duch", plan.Warnings[0]);
        Assert.Single(plan.Files);
        Assert.True(plan.Files[0].IsGroup);
        Assert.Equal("Menadzer", plan.Files[0].Name);
    }

    [Fact]
    public void Grupa_bez_nazwy_albo_bez_wartosci_blokuje_start()
    {
        var config = Config();
        config.Groups =
        [
            new RecipientGroup { Name = "", Values = ["Jan Kowalski"] },
            new RecipientGroup { Name = "Pusta", Values = [] },
        ];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Equal(2, plan.Problems.Count);
        Assert.Empty(plan.Files);
    }

    [Fact]
    public void Wartosc_w_dwoch_grupach_daje_ostrzezenie()
    {
        var config = Config();
        config.Groups =
        [
            new RecipientGroup { Name = "Pierwszy", Values = ["Jan Kowalski"] },
            new RecipientGroup { Name = "Drugi", Values = ["Jan Kowalski"] },
        ];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Empty(plan.Problems);
        Assert.Single(plan.Warnings);
        Assert.Contains("Jan Kowalski", plan.Warnings[0]);
    }

    [Fact]
    public void Dwa_pliki_o_tej_samej_nazwie_blokuja_start()
    {
        var config = Config();
        config.OutputFolder = Path.Combine(Path.GetTempPath(), "efs-test-out");
        config.Groups =
        [
            new RecipientGroup { Name = "A/B", Values = ["Jan Kowalski"] },
            new RecipientGroup { Name = "A_B", Values = ["Anna Nowak"] },
        ];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski", "Anna Nowak"));

        Assert.Contains(plan.Problems, problem => problem.Contains("ten sam plik"));
    }

    [Fact]
    public void Plik_bez_wlasnych_arkuszy_dostaje_globalna_liste()
    {
        var config = Config();
        config.SheetsToKeep = ["Dane", "Dashboard"];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Equal(new[] { "Dane", "Dashboard" }, plan.Files[0].SheetsToKeep);
    }

    [Fact]
    public void Grupa_doklada_swoje_arkusze_do_globalnej_bazy()
    {
        var config = Config();
        config.SheetsToKeep = ["Dane"];
        config.Groups =
        [
            new RecipientGroup { Name = "Polnoc", Values = ["Jan Kowalski"], SheetsToKeep = ["Raport 1", "Raport 2"] },
            new RecipientGroup { Name = "Poludnie", Values = ["Anna Nowak"], SheetsToKeep = ["Raport 2"] },
        ];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski", "Anna Nowak"));

        Assert.Equal(new[] { "Dane", "Raport 1", "Raport 2" }, plan.Files[0].SheetsToKeep);
        Assert.Equal(new[] { "Dane", "Raport 2" }, plan.Files[1].SheetsToKeep);
    }

    [Fact]
    public void Arkusz_wpisany_i_globalnie_i_przy_grupie_nie_dubluje_sie()
    {
        var config = Config();
        config.SheetsToKeep = ["Dane", "Dashboard"];
        config.Groups = [new RecipientGroup { Name = "Polnoc", Values = ["Jan Kowalski"], SheetsToKeep = ["dashboard", "Raport 1"] }];

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Jan Kowalski"));

        Assert.Equal(new[] { "Dane", "Dashboard", "Raport 1" }, plan.Files[0].SheetsToKeep);
    }

    [Fact]
    public void Wartosc_z_nowa_linia_blokuje_start()
    {
        RecipientPlan plan = Recipients.BuildPlan(Config(), Klucze("Jan\nKowalski"));

        Assert.Contains(plan.Problems, problem => problem.Contains("nowej linii"));
    }

    private static SplitConfig MultiColumnConfig() => new()
    {
        Mode = SplitMode.Worksheet,
        SourceWorkbook = @"C:\raporty\Sprzedaz.xlsx",
        SplitColumns = ["Region", "Handlowiec"],
        FileNameTemplate = "Raport_{value}.xlsx",
    };

    [Fact]
    public void Plik_powstaje_na_kazda_kombinacje_z_danych()
    {
        RecipientPlan plan = Recipients.BuildPlan(
            MultiColumnConfig(),
            [SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("Poludnie", "Nowak")]);

        Assert.Equal(2, plan.Files.Count);
        Assert.Empty(plan.Problems);
        Assert.Equal("Polnoc - Kowalski", plan.Files[0].Name);
        Assert.Equal(SplitKey.Of("Polnoc", "Kowalski"), Assert.Single(plan.Files[0].Keys));
    }

    [Fact]
    public void Te_same_wartosci_w_innych_parach_to_osobne_pliki()
    {
        RecipientPlan plan = Recipients.BuildPlan(
            MultiColumnConfig(),
            [SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("Poludnie", "Kowalski")]);

        Assert.Equal(2, plan.Files.Count);
        Assert.Equal(new[] { "Polnoc - Kowalski", "Poludnie - Kowalski" }, plan.Files.Select(file => file.Name));
    }

    [Fact]
    public void Kombinacja_z_konfiguracji_dostaje_pisownie_z_danych()
    {
        var config = MultiColumnConfig();
        config.Keys = [["polnoc", "kowalski"]];

        RecipientPlan plan = Recipients.BuildPlan(config, [SplitKey.Of("POLNOC", "Kowalski")]);

        Assert.Equal("POLNOC - Kowalski", Assert.Single(plan.Files).Name);
    }

    [Fact]
    public void Kombinacji_spoza_danych_nie_da_sie_wygenerowac()
    {
        var config = MultiColumnConfig();
        config.Keys = [["Polnoc", "Nowak"]];

        RecipientPlan plan = Recipients.BuildPlan(
            config, [SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("Poludnie", "Nowak")]);

        Assert.Contains(plan.Problems, problem => problem.Contains("Polnoc - Nowak"));
    }

    [Fact]
    public void Grupa_zbiera_kilka_kombinacji_w_jeden_plik()
    {
        var config = MultiColumnConfig();
        config.SheetsToKeep = ["Dashboard"];
        config.Groups =
        [
            new RecipientGroup { Name = "Zarzad", Keys = [["Polnoc", "Kowalski"], ["Poludnie", "Nowak"]] },
        ];

        RecipientPlan plan = Recipients.BuildPlan(
            config, [SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("Poludnie", "Nowak")]);

        Recipient zarzad = Assert.Single(plan.Files);
        Assert.Equal("Zarzad", zarzad.Name);
        Assert.Equal(2, zarzad.Keys.Count);
        Assert.True(zarzad.IsGroup);
    }

    [Fact]
    public void Plik_odbiorcy_nie_moze_nadpisac_skoroszytu_zrodlowego()
    {
        var config = Config();
        config.SourceWorkbook = @"C:\raporty\Raport_Anna Nowak.xlsx";
        config.OutputFolder = @"C:\raporty";
        config.FileNameTemplate = "Raport_{value}.xlsx";

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Anna Nowak", "Jan Kowalski"));

        string problem = Assert.Single(plan.Problems);
        Assert.Contains("Anna Nowak", problem);
        Assert.Contains("skoroszyt źródłowy", problem);
    }

    [Fact]
    public void Plik_odbiorcy_nie_moze_nadpisac_skoroszytu_przygotowanego()
    {
        var config = Config();
        config.SourceWorkbook = @"C:\raporty\Raport.xlsm";
        config.OutputFolder = @"C:\raporty";
        config.FileNameTemplate = "Raport.prepared.xlsm";

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Anna Nowak"));

        string problem = Assert.Single(plan.Problems);
        Assert.Contains("przygotowany", problem);
    }

    [Fact]
    public void W_trybie_Worksheet_nazwa_prepared_nie_jest_problemem()
    {
        var config = Config();
        config.Mode = SplitMode.Worksheet;
        config.SourceWorkbook = @"C:\raporty\Raport.xlsm";
        config.OutputFolder = @"C:\raporty";
        config.FileNameTemplate = "Raport.prepared.xlsm";

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Anna Nowak"));

        Assert.Empty(plan.Problems);
    }

    [Fact]
    public void Pliki_obok_zrodla_ale_o_innej_nazwie_sa_w_porzadku()
    {
        var config = Config();
        config.SourceWorkbook = @"C:\raporty\Raport.xlsm";
        config.OutputFolder = @"C:\raporty";

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Anna Nowak", "Jan Kowalski"));

        Assert.Empty(plan.Problems);
    }

    [Fact]
    public void Niepoprawna_sciezka_zrodla_nie_wywraca_planu()
    {
        var config = Config();
        config.SourceWorkbook = "C:\\raporty\\Raport\0.xlsm";
        config.OutputFolder = @"C:\raporty\out";

        RecipientPlan plan = Recipients.BuildPlan(config, Klucze("Anna Nowak"));

        Assert.Single(plan.Files);
        Assert.Empty(plan.Problems);
    }
}
