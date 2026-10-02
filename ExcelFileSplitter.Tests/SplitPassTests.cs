using System.Reflection;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class SplitPassTests : IDisposable
{
    private readonly string _sourceWorkbook;

    public SplitPassTests()
    {
        _sourceWorkbook = Path.Combine(Path.GetTempPath(), $"efs-pass-{Guid.NewGuid():N}.xlsb");
        File.WriteAllText(_sourceWorkbook, "");
    }

    private SplitConfig TwoPasses() => new()
    {
        Mode = SplitMode.Worksheet,
        SourceWorkbook = _sourceWorkbook,
        OutputFolder = Path.Combine(Path.GetTempPath(), "efs-pass-out"),
        Worksheet = new WorksheetOptions { Sheet = "Baza", HeaderRow = 1 },
        SheetsToKeep = ["Dashboard"],
        FileNameTemplate = "Raport_{podzial}_{value}.xlsx",
        Passes =
        [
            new SplitPass { Name = "Menadzer", SplitColumn = "Menadzer", SheetsToKeep = ["Raport menadzera"] },
            new SplitPass { Name = "Obsluga", SplitColumn = "Obsluga klienta" },
        ],
    };

    [Fact]
    public void Bez_sekcji_Passes_powstaje_jeden_podzial_z_pol_na_gorze()
    {
        var config = new SplitConfig
        {
            SplitColumn = "PH",
            Values = ["ANNA", "PAWEL"],
            Groups = [new RecipientGroup { Name = "Polnoc", Values = ["ANNA"] }],
        };

        IReadOnlyList<SplitPass> passes = config.SplitPasses();

        Assert.Single(passes);
        Assert.Equal("PH", passes[0].SplitColumn);
        Assert.Same(config.Values, passes[0].Values);
        Assert.Same(config.Groups, passes[0].Groups);
    }

    [Fact]
    public void Podzial_dostaje_wlasne_kolumny_wartosci_i_grupy()
    {
        SplitConfig config = TwoPasses();
        config.Passes[1].Values = ["ANNA NOWAK"];

        SplitConfig second = config.ForPass(config.Passes[1]);

        Assert.Equal("Obsluga klienta", second.SplitColumn);
        Assert.Single(second.Values);
        Assert.Equal("ANNA NOWAK", second.Values[0]);
        Assert.Empty(second.Passes);
    }

    [Fact]
    public void Arkusze_podzialu_dokladaja_sie_do_wspolnych_a_nie_zastepuja_ich()
    {
        SplitConfig config = TwoPasses();

        SplitConfig first = config.ForPass(config.Passes[0]);
        SplitConfig second = config.ForPass(config.Passes[1]);

        Assert.Equal("Dashboard, Raport menadzera", string.Join(", ", first.SheetsToKeep));
        Assert.Equal("Dashboard", string.Join(", ", second.SheetsToKeep));
    }

    [Fact]
    public void Nazwa_podzialu_trafia_w_nazwe_pliku()
    {
        SplitConfig config = TwoPasses();

        string manager = config.ForPass(config.Passes[0]).OutputFileNameFor("JAN KOWALSKI");
        string service = config.ForPass(config.Passes[1]).OutputFileNameFor("JAN KOWALSKI");

        Assert.Equal("Raport_Menadzer_JAN KOWALSKI.xlsb", manager);
        Assert.Equal("Raport_Obsluga_JAN KOWALSKI.xlsb", service);
        Assert.NotEqual(manager, service);
    }

    [Fact]
    public void Bez_nazwy_podzialu_znacznik_znika_razem_z_separatorem()
    {
        Assert.Equal("Raport_{value}.xlsx", SplitConfig.ResolvePassMarker("Raport_{podzial}_{value}.xlsx", ""));
        Assert.Equal("Raport_{value}.xlsx", SplitConfig.ResolvePassMarker("Raport_{value}_{podzial}.xlsx", ""));
        Assert.Equal("Raport_Menadzer_{value}.xlsx", SplitConfig.ResolvePassMarker("Raport_{podzial}_{value}.xlsx", "Menadzer"));
    }

    [Fact]
    public void Znacznik_pass_dziala_tak_samo_jak_podzial()
    {
        Assert.Equal("Raport_Menadzer_{value}.xlsx", SplitConfig.ResolvePassMarker("Raport_{pass}_{value}.xlsx", "Menadzer"));
    }

    [Fact]
    public void Nie_wolno_opisac_podzialu_na_dwa_sposoby_naraz()
    {
        SplitConfig config = TwoPasses();
        config.SplitColumn = "PH";

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Passes", ex.Message);
    }

    [Fact]
    public void Kilka_podzialow_dziala_tylko_w_trybie_Worksheet()
    {
        SplitConfig config = TwoPasses();
        config.Mode = SplitMode.PowerPivot;

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Worksheet", ex.Message);
    }

    [Fact]
    public void Kazdy_podzial_musi_miec_nazwe()
    {
        SplitConfig config = TwoPasses();
        config.Passes[1].Name = "";

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Name", ex.Message);
    }

    [Fact]
    public void Nazwy_podzialow_musza_sie_roznic()
    {
        SplitConfig config = TwoPasses();
        config.Passes[1].Name = "menadzer";

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("więcej niż raz", ex.Message);
    }

    [Fact]
    public void Blad_w_jednym_podziale_zatrzymuje_calosc()
    {
        SplitConfig config = TwoPasses();
        config.Passes[1].SplitColumn = "";

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("SplitColumn", ex.Message);
    }

    [Fact]
    public void Poprawna_konfiguracja_z_dwoma_podzialami_przechodzi()
    {
        TwoPasses().Validate();
    }

    [Fact]
    public void Podzialy_przezywaja_zapis_i_odczyt()
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-pass-cfg-{Guid.NewGuid():N}.json");
        try
        {
            TwoPasses().Save(path);
            SplitConfig loaded = SplitConfig.Load(path);

            Assert.Equal(2, loaded.Passes.Count);
            Assert.Equal("Menadzer", loaded.Passes[0].Name);
            Assert.Equal("Obsluga klienta", loaded.Passes[1].SplitColumn);
            Assert.Equal("Raport menadzera", string.Join(", ", loaded.Passes[0].SheetsToKeep));
            loaded.Validate();
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void Kazde_pole_podzialu_trafia_do_konfiguracji()
    {
        var pass = new SplitPass
        {
            Name = "Kontrola",
            SplitColumn = "Kolumna podzialu",
            SplitColumns = ["Pierwsza", "Druga"],
            Values = ["WARTOSC A"],
            Keys = [["Pierwsza A", "Druga A"]],
            Groups = [new RecipientGroup { Name = "Grupa", Values = ["WARTOSC A"] }],
            SheetsToKeep = ["Arkusz podzialu"],
            FileNameTemplate = "Wlasny_{podzial}_{value}.xlsx",
            OutputFolder = Path.Combine(Path.GetTempPath(), "efs-pass-wlasny"),
        };

        SplitConfig config = TwoPasses();
        SplitConfig forPass = config.ForPass(pass);

        var checks = new Dictionary<string, Action>(StringComparer.Ordinal)
        {
            [nameof(SplitPass.Name)] =
                () => Assert.Contains("Kontrola", forPass.FileNameTemplate),

            [nameof(SplitPass.SplitColumn)] =
                () => Assert.Equal(pass.SplitColumn, forPass.SplitColumn),

            [nameof(SplitPass.SplitColumns)] =
                () => Assert.Equal(pass.SplitColumns, forPass.SplitColumns),

            [nameof(SplitPass.Values)] =
                () => Assert.Equal(pass.Values, forPass.Values),

            [nameof(SplitPass.Keys)] =
                () => Assert.Equal(pass.Keys, forPass.Keys),

            [nameof(SplitPass.Groups)] =
                () => Assert.Same(pass.Groups, forPass.Groups),

            [nameof(SplitPass.SheetsToKeep)] =
                () => Assert.Equal("Dashboard, Arkusz podzialu", string.Join(", ", forPass.SheetsToKeep)),

            [nameof(SplitPass.FileNameTemplate)] =
                () => Assert.StartsWith("Wlasny_", forPass.FileNameTemplate),

            [nameof(SplitPass.OutputFolder)] =
                () => Assert.Equal(pass.OutputFolder, forPass.OutputFolder),

            [nameof(SplitPass.SplitColumnNames)] =
                () => Assert.Equal(pass.SplitColumnNames, forPass.SplitColumnNames),
        };

        var withoutCheck = new List<string>();

        foreach (PropertyInfo property in typeof(SplitPass).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (checks.TryGetValue(property.Name, out Action? check)) check();
            else withoutCheck.Add(property.Name);
        }

        Assert.True(withoutCheck.Count == 0,
            $"SplitPass ma pole bez sprawdzenia: {string.Join(", ", withoutCheck)}. " +
            "Dopisz je do tablicy `checks` w tym tescie ORAZ upewnij sie, ze SplitConfig.ForPass " +
            "faktycznie je przenosi - inaczej ustawienie uzytkownika bedzie po cichu ignorowane.");

        var namesOfPass = typeof(SplitPass)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Select(property => property.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.All(checks.Keys, name =>
            Assert.True(namesOfPass.Contains(name), $"sprawdzenie dotyczy pola '{name}', ktorego SplitPass juz nie ma"));
    }

    [Fact]
    public void Konfiguracja_bez_sekcji_Passes_nie_traci_zadnego_ustawienia()
    {
        var config = new SplitConfig
        {
            Mode = SplitMode.Worksheet,
            SourceWorkbook = _sourceWorkbook,
            OutputFolder = Path.Combine(Path.GetTempPath(), "efs-bez-passes"),
            Worksheet = new WorksheetOptions { Sheet = "Baza", HeaderRow = 1 },
            SplitColumn = "PH",
            SplitColumns = ["PH", "Region"],
            Values = ["ANNA"],
            Keys = [["ANNA", "Polnoc"]],
            Groups = [new RecipientGroup { Name = "Polnoc", Values = ["ANNA"] }],
            SheetsToKeep = ["Dashboard", "Okladka"],
            FileNameTemplate = "Raport_{value}.xlsx",
        };

        SplitConfig single = config.ForPass(config.SplitPasses()[0]);

        Assert.Equal(config.SplitColumn, single.SplitColumn);
        Assert.Equal(config.SplitColumns, single.SplitColumns);
        Assert.Equal(config.Values, single.Values);
        Assert.Equal(config.Keys, single.Keys);
        Assert.Same(config.Groups, single.Groups);
        Assert.Equal("Dashboard, Okladka", string.Join(", ", single.SheetsToKeep));
        Assert.Equal(config.FileNameTemplate, single.FileNameTemplate);
        Assert.Equal(config.OutputFolder, single.OutputFolder);
        Assert.Equal(config.SplitColumnNames, single.SplitColumnNames);
    }

    public void Dispose()
    {
        try { File.Delete(_sourceWorkbook); } catch { }
    }
}
