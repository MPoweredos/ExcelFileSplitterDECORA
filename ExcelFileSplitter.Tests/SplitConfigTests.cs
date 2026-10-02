using ExcelFileSplitter.Core;
using ExcelFileSplitter.Interop;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class SplitConfigTests : IDisposable
{
    private readonly string _sourceWorkbook;

    public SplitConfigTests()
    {
        _sourceWorkbook = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.xlsm");
        File.WriteAllText(_sourceWorkbook, "");
    }

    private SplitConfig ValidConfig() => new()
    {
        SourceWorkbook = _sourceWorkbook,
        OutputFolder = Path.Combine(Path.GetTempPath(), "efs-test-out"),
        SplitColumn = "Opiekun",
        SheetsToKeep = ["Dashboard"],
        FileNameTemplate = "Raport_{value}.xlsx",
        PowerPivot = new PowerPivotOptions { SplitTable = "dim_Klienci" },
    };

    [Fact]
    public void Rozszerzenie_pliku_wynikowego_idzie_za_zrodlem_a_nie_za_szablonem()
    {
        var config = ValidConfig();

        string path = config.OutputPathFor("Jan Kowalski");

        Assert.EndsWith(".xlsm", path);
        Assert.Contains("Raport_Jan Kowalski", path);
    }

    [Fact]
    public void Znacznik_name_dziala_tak_samo_jak_value()
    {
        var config = ValidConfig();
        config.FileNameTemplate = "Raport_{name}.xlsx";

        Assert.Contains("Raport_Anna Nowak", config.OutputPathFor("Anna Nowak"));
    }

    [Fact]
    public void Znaki_niedozwolone_w_nazwie_pliku_sa_zamieniane()
    {
        var config = ValidConfig();

        string fileName = Path.GetFileName(config.OutputPathFor("Region A/B"));

        Assert.Equal("Raport_Region A_B.xlsm", fileName);
    }

    [Fact]
    public void Poprawna_konfiguracja_przechodzi_walidacje()
    {
        ValidConfig().Validate();
    }

    [Fact]
    public void Values_i_Groups_naraz_sa_odrzucane()
    {
        var config = ValidConfig();
        config.Values = ["Jan Kowalski"];
        config.Groups = [new RecipientGroup { Name = "Menadzer", Values = ["Anna Nowak"] }];

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Values", ex.Message);
        Assert.Contains("Groups", ex.Message);
    }

    [Fact]
    public void Filtr_wymiaru_na_innej_kolumnie_niz_SplitColumn_jest_odrzucany()
    {
        var config = ValidConfig();
        config.PowerPivot.Queries = [new QuerySpec { Name = "dim_Klienci", FilterColumn = "Region", Cache = true }];

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Region", ex.Message);
        Assert.Contains("Opiekun", ex.Message);
    }

    [Fact]
    public void Filtr_faktow_przez_klucz_moze_byc_na_innej_kolumnie()
    {
        var config = ValidConfig();
        config.PowerPivot.Queries =
        [
            new QuerySpec
            {
                Name = "dim_Klienci",
                FilterColumn = "Odbiorca",
                KeyFrom = new KeyFrom { Query = "dim_Klienci", Column = "Identyfikator" },
            },
        ];

        config.Validate();
    }

    [Fact]
    public void Szablon_nazwy_bez_znacznika_jest_odrzucany()
    {
        var config = ValidConfig();
        config.FileNameTemplate = "Raport.xlsx";

        Assert.Throws<InvalidOperationException>(config.Validate);
    }

    [Fact]
    public void UseCache_bez_CacheFolder_jest_odrzucany()
    {
        var config = ValidConfig();
        config.PowerPivot.UseCache = true;
        config.PowerPivot.CacheFolder = "";

        Assert.Throws<InvalidOperationException>(config.Validate);
    }

    [Fact]
    public void Pusta_lista_arkuszy_jest_odrzucana()
    {
        var config = ValidConfig();
        config.SheetsToKeep = [];

        Assert.Throws<InvalidOperationException>(config.Validate);
    }

    [Fact]
    public void Brak_skoroszytu_zrodlowego_jest_wykrywany()
    {
        var config = ValidConfig();
        config.SourceWorkbook = Path.Combine(Path.GetTempPath(), $"efs-nie-ma-{Guid.NewGuid():N}.xlsm");

        Assert.Throws<FileNotFoundException>(config.Validate);
    }

    [Fact]
    public void Sciezka_przygotowanego_skoroszytu_powstaje_obok_zrodla()
    {
        var config = ValidConfig();
        config.PowerPivot.PreparedWorkbook = "";

        string prepared = config.PreparedWorkbookPath();

        Assert.EndsWith(".prepared.xlsm", prepared);
        Assert.Equal(Path.GetDirectoryName(_sourceWorkbook), Path.GetDirectoryName(prepared));
    }

    private SplitConfig ValidWorksheetConfig() => new()
    {
        Mode = SplitMode.Worksheet,
        SourceWorkbook = _sourceWorkbook,
        OutputFolder = Path.Combine(Path.GetTempPath(), "efs-test-out"),
        Worksheet = new WorksheetOptions { Sheet = "Dane", HeaderRow = 1 },
        SplitColumn = "Opiekun",
        SheetsToKeep = ["Dane", "Dashboard"],
        FileNameTemplate = "Raport_{value}.xlsx",
    };

    [Fact]
    public void Tryb_Worksheet_nie_wymaga_SplitTable()
    {
        var config = ValidWorksheetConfig();
        config.PowerPivot.SplitTable = "";

        config.Validate();
    }

    [Fact]
    public void Tryb_Worksheet_bez_wskazania_danych_jest_odrzucany()
    {
        var config = ValidWorksheetConfig();
        config.Worksheet = new WorksheetOptions();

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("Worksheet.Table", ex.Message);
        Assert.Contains("Worksheet.Sheet", ex.Message);
    }

    [Fact]
    public void Arkusza_z_danymi_nie_trzeba_zostawiac_u_odbiorcy()
    {
        var config = ValidWorksheetConfig();
        config.SheetsToKeep = ["Dashboard"];

        config.Validate();
    }

    [Fact]
    public void Tabela_Excela_zwalnia_z_podawania_arkusza()
    {
        var config = ValidWorksheetConfig();
        config.Worksheet = new WorksheetOptions { Table = "tblSprzedaz" };

        config.Validate();
    }

    [Fact]
    public void Wiersz_naglowkow_mniejszy_od_jeden_jest_odrzucany()
    {
        var config = ValidWorksheetConfig();
        config.Worksheet.HeaderRow = 0;

        Assert.Throws<InvalidOperationException>(config.Validate);
    }

    [Fact]
    public void UseCache_w_trybie_Worksheet_jest_odrzucany()
    {
        var config = ValidWorksheetConfig();
        config.PowerPivot.UseCache = true;
        config.PowerPivot.CacheFolder = Path.GetTempPath();

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("PowerPivot", ex.Message);
    }

    [Fact]
    public void Queries_w_trybie_Worksheet_sa_odrzucane()
    {
        var config = ValidWorksheetConfig();
        config.PowerPivot.Queries = [new QuerySpec { Name = "dim_Klienci" }];

        var ex = Assert.Throws<InvalidOperationException>(config.Validate);
        Assert.Contains("PowerPivot", ex.Message);
    }

    [Fact]
    public void Tryb_Worksheet_pracuje_na_oryginale_nawet_gdy_obok_lezy_plik_prepared()
    {
        var config = ValidWorksheetConfig();
        string prepared = config.PreparedWorkbookPath();
        File.WriteAllText(prepared, "");
        try
        {
            Assert.Equal(Path.GetFullPath(_sourceWorkbook), config.WorkbookToProcess());
        }
        finally
        {
            try { File.Delete(prepared); } catch { }
        }
    }

    [Fact]
    public void Tryb_zapisuje_sie_w_json_jako_nazwa_a_nie_liczba()
    {
        var config = ValidWorksheetConfig();
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-cfg-{Guid.NewGuid():N}.json");
        try
        {
            config.Save(path);
            Assert.Contains("\"Worksheet\"", File.ReadAllText(path));
            Assert.Equal(SplitMode.Worksheet, SplitConfig.Load(path).Mode);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void Nazwa_pliku_bez_folderu_da_sie_pokazac_zanim_folder_zostanie_wskazany()
    {
        var config = ValidConfig();
        config.OutputFolder = "";

        Assert.Equal("Raport_Jan Kowalski.xlsm", config.OutputFileNameFor("Jan Kowalski"));
    }

    [Fact]
    public void Plik_ustawien_znosi_komentarze_i_przecinek_na_koncu()
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-cfg-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """
                // Ustawienia podzialu - komentarz na gorze pliku.
                {
                  "Mode": "Worksheet",       // komentarz na koncu linii
                  "SplitColumn": "Opiekun",
                  "SheetsToKeep": [
                    "Dashboard",
                  ],
                }
                """);

            SplitConfig config = SplitConfig.Load(path);

            Assert.Equal(SplitMode.Worksheet, config.Mode);
            Assert.Equal("Opiekun", config.SplitColumn);
            Assert.Single(config.SheetsToKeep);
            Assert.Equal("Dashboard", config.SheetsToKeep[0]);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void Szablon_ustawien_tlumaczy_Values_Groups_i_cache_i_daje_sie_wczytac()
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-cfg-{Guid.NewGuid():N}.json");
        try
        {
            SplitConfig.CreateSample().SaveTemplate(path);
            string text = File.ReadAllText(path);

            Assert.Contains("\"Values\"", text);
            Assert.Contains("\"Groups\"", text);
            Assert.Contains("SheetsToKeep", text);
            Assert.Contains("UseCache", text);
            Assert.Contains("Queries[].Cache", text);
            Assert.StartsWith("//", text.TrimStart());

            Assert.Equal(SplitMode.PowerPivot, SplitConfig.Load(path).Mode);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void Zamiana_formul_na_wartosci_jest_domyslnie_wlaczona_i_da_sie_ja_zapisac()
    {
        Assert.True(new WorksheetOptions().FormulasToValues);

        var config = ValidWorksheetConfig();
        config.Worksheet.FormulasToValues = false;

        string path = Path.Combine(Path.GetTempPath(), $"efs-test-cfg-{Guid.NewGuid():N}.json");
        try
        {
            config.Save(path);
            Assert.False(SplitConfig.Load(path).Worksheet.FormulasToValues);
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void Puste_wiersze_za_tabela_sa_policzone_osobno()
    {
        var withJunk = new DataRange("Baza", null, 1, 2, 10_769, 1, 20, ["PH"]) { SheetLastRow = 102_613 };
        var clean = new DataRange("Baza", null, 1, 2, 10_769, 1, 20, ["PH"]) { SheetLastRow = 10_769 };
        var unknown = new DataRange("Baza", "tblDane", 1, 2, 10_769, 1, 20, ["PH"]);

        Assert.Equal(91_844, withJunk.EmptyRowsAfterData);
        Assert.Equal(0, clean.EmptyRowsAfterData);

        Assert.Equal(0, unknown.EmptyRowsAfterData);
    }

    [Fact]
    public void Przykladowa_konfiguracja_zwyklego_pliku_jest_spojna()
    {
        SplitConfig sample = SplitConfig.CreateWorksheetSample();

        Assert.Equal(SplitMode.Worksheet, sample.Mode);
        Assert.False(sample.PowerPivot.UseCache);
        Assert.Empty(sample.PowerPivot.Queries);
        Assert.NotEqual("", sample.Worksheet.Table);
    }

    public void Dispose()
    {
        try { File.Delete(_sourceWorkbook); } catch { }
    }

    private SplitConfig MultiColumnConfig()
    {
        var config = ValidConfig();
        config.Mode = SplitMode.Worksheet;
        config.PowerPivot.SplitTable = "";
        config.SplitColumn = "";
        config.SplitColumns = ["Region", "Handlowiec"];
        config.Worksheet = new WorksheetOptions { Sheet = "Baza", HeaderRow = 1 };
        return config;
    }

    [Fact]
    public void Podzial_po_dwoch_kolumnach_z_kombinacjami_przechodzi()
    {
        var config = MultiColumnConfig();
        config.Keys = [["Polnoc", "Kowalski"], ["Poludnie", "Nowak"]];

        config.Validate();

        Assert.Equal(new[] { "Region", "Handlowiec" }, config.SplitColumnNames);
    }

    [Fact]
    public void Nie_da_sie_ustawic_SplitColumn_i_SplitColumns_naraz()
    {
        var config = MultiColumnConfig();
        config.SplitColumn = "Opiekun";

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("SplitColumns", error.Message);
    }

    [Fact]
    public void Ta_sama_kolumna_dwa_razy_jest_bledem()
    {
        var config = MultiColumnConfig();
        config.SplitColumns = ["Region", "region"];

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("więcej niż raz", error.Message);
    }

    [Fact]
    public void Przy_kilku_kolumnach_Values_nie_wystarcza()
    {
        var config = MultiColumnConfig();
        config.Values = ["Polnoc"];

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("Keys", error.Message);
    }

    [Fact]
    public void Kombinacja_o_zlej_liczbie_wartosci_jest_bledem()
    {
        var config = MultiColumnConfig();
        config.Keys = [["Polnoc", "Kowalski", "za duzo"]];

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("kolumn podziału jest 2", error.Message);
    }

    [Fact]
    public void Grupa_przy_kilku_kolumnach_musi_uzywac_Keys()
    {
        var config = MultiColumnConfig();
        config.Groups = [new RecipientGroup { Name = "Zarzad", Values = ["Polnoc"] }];

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("Zarzad", error.Message);
    }

    [Fact]
    public void Tryb_PowerPivot_odmawia_podzialu_po_kilku_kolumnach()
    {
        var config = ValidConfig();
        config.SplitColumn = "";
        config.SplitColumns = ["Region", "Handlowiec"];

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("Worksheet", error.Message);
    }

    [Fact]
    public void Brak_kolumny_podzialu_jest_bledem()
    {
        var config = ValidConfig();
        config.SplitColumn = "";

        var error = Assert.Throws<InvalidOperationException>(config.Validate);

        Assert.Contains("SplitColumn", error.Message);
    }

    [Fact]
    public void Kazde_wskazanie_kombinacji_czyta_wartosci_tak_samo()
    {
        List<IKeySelection> wskazania =
        [
            new SplitConfig { Values = ["ANNA", "PAWEL"] },
            new SplitPass { Values = ["ANNA", "PAWEL"] },
            new RecipientGroup { Values = ["ANNA", "PAWEL"] },
        ];

        foreach (IKeySelection wskazanie in wskazania)
        {
            List<SplitKey> keys = wskazanie.RequestedKeys();

            Assert.Equal(2, keys.Count);
            Assert.Equal("ANNA", keys[0].Display);
            Assert.Equal("PAWEL", keys[1].Display);
        }
    }

    [Fact]
    public void Keys_maja_pierwszenstwo_przed_Values()
    {
        List<IKeySelection> wskazania =
        [
            new SplitConfig { Values = ["PAWEL"], Keys = [["ANNA", "Polnoc"]] },
            new SplitPass { Values = ["PAWEL"], Keys = [["ANNA", "Polnoc"]] },
            new RecipientGroup { Values = ["PAWEL"], Keys = [["ANNA", "Polnoc"]] },
        ];

        foreach (IKeySelection wskazanie in wskazania)
        {
            List<SplitKey> keys = wskazanie.RequestedKeys();

            Assert.Single(keys);
            Assert.Equal("ANNA - Polnoc", keys[0].Display);
        }
    }

    [Fact]
    public void Przyklad_zwyklego_pliku_zostawia_sekcje_PowerPivot_domyslna()
    {
        SplitConfig worksheet = SplitConfig.CreateWorksheetSample();
        var untouched = new PowerPivotOptions();

        Assert.Equal(untouched.CacheFolder, worksheet.PowerPivot.CacheFolder);
        Assert.Equal(untouched.SplitTable, worksheet.PowerPivot.SplitTable);
        Assert.Empty(worksheet.PowerPivot.Queries);
        Assert.Empty(worksheet.PowerPivot.CacheTables);
    }

    [Fact]
    public void Plik_zrodlowy_musi_byc_skoroszytem_Excela()
    {
        string notExcel = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.txt");
        File.WriteAllText(notExcel, "to nie jest skoroszyt");
        try
        {
            var config = ValidConfig();
            config.SourceWorkbook = notExcel;

            var error = Assert.Throws<InvalidOperationException>(config.Validate);

            Assert.Contains("nie jest plikiem Excela", error.Message);
            Assert.Contains("xlsx", error.Message);
        }
        finally
        {
            try { File.Delete(notExcel); } catch { }
        }
    }

    [Fact]
    public void Tryb_PowerPivot_odrzuca_stary_format_xls()
    {
        string old = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.xls");
        File.WriteAllText(old, "");
        try
        {
            var config = ValidConfig();
            config.SourceWorkbook = old;

            var error = Assert.Throws<InvalidOperationException>(config.Validate);

            Assert.Contains("modelu danych", error.Message);
        }
        finally
        {
            try { File.Delete(old); } catch { }
        }
    }
}
