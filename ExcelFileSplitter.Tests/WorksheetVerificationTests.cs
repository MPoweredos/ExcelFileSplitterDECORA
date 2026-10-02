using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class WorksheetVerificationTests
{
    private static SplitKey Klucz(params string[] values) => SplitKey.Of(values);

    private static WorksheetVerificationResult Clean() => new(
        FilePath: @"C:\out\Raport_Jan.xlsx",
        ExpectedKeys: [Klucz("Jan Kowalski")],
        KeysInData: [Klucz("Jan Kowalski")],
        KeptRowCount: 120,
        DataSheetInFile: true,
        StaleSlicerItems: [],
        BrokenFormulas: [],
        Leaks: [],
        NameCoincidences: [],
        Error: null);

    [Fact]
    public void Czysty_plik_przechodzi()
    {
        WorksheetVerificationResult result = Clean();

        Assert.True(result.Ok);
        Assert.Contains("OK", result.Summary());
        Assert.Contains("120", result.Summary());
    }

    [Fact]
    public void Obca_wartosc_w_kolumnie_podzialu_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with { KeysInData = [Klucz("Jan Kowalski"), Klucz("Anna Nowak")] };

        Assert.False(result.Ok);
        Assert.Equal(new[] { "Anna Nowak" }, result.ForeignKeys.Select(key => key.Display));
        Assert.Contains("Anna Nowak", result.Summary());
    }

    [Fact]
    public void Pusty_plik_blokuje_wyslanie()
    {
        WorksheetVerificationResult result = Clean() with { KeysInData = [], KeptRowCount = 0 };

        Assert.False(result.Ok);
        Assert.Contains("nie został żaden wiersz", result.Summary());
    }

    [Fact]
    public void Tabela_przestawna_na_cache_sprzed_podzialu_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with
        {
            StalePivotCaches = ["'Tabela przestawna2' (arkusz 'Raport') - cache z 2026-09-01 10:00:00"],
        };

        Assert.False(result.Ok);
        Assert.Contains("Tabela przestawna2", result.Summary());
        Assert.Contains("sprzed podziału", result.Summary());
    }

    [Fact]
    public void Wyciek_w_paczce_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with
        {
            Leaks = ["xl/charts/chart1.xml -> Anna Nowak"],
        };

        Assert.False(result.Ok);
        Assert.Contains("WYCIEK", result.Summary());
        Assert.Contains("chart1.xml", result.Summary());
    }

    [Fact]
    public void Formula_z_REF_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with { BrokenFormulas = ["Dashboard: =SUMA(Dane!#REF!)"] };

        Assert.False(result.Ok);
        Assert.Contains("#REF!", result.Summary());
    }

    [Fact]
    public void Fragmentator_z_obcym_elementem_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with { StaleSlicerItems = ["Opiekun[Opiekun] -> Anna Nowak"] };

        Assert.False(result.Ok);
        Assert.Contains("fragmentatory", result.Summary());
    }

    [Fact]
    public void Blad_weryfikacji_nie_jest_zgoda_na_wyslanie()
    {
        WorksheetVerificationResult result = Clean() with { Error = "nie udalo sie otworzyc pliku" };

        Assert.False(result.Ok);
        Assert.StartsWith("BŁĄD WERYFIKACJI", result.Summary());
    }

    [Fact]
    public void Zbieg_nazw_nie_blokuje_ale_jest_odnotowany()
    {
        WorksheetVerificationResult result = Clean() with { NameCoincidences = ["Nieprzypisane"] };

        Assert.True(result.Ok);
        Assert.Contains("zbiegi nazw", result.Summary());
    }

    [Fact]
    public void Wartosc_odbiorcy_bez_danych_to_informacja_a_nie_blad()
    {
        WorksheetVerificationResult result = Clean() with
        {
            ExpectedKeys = [Klucz("Jan Kowalski"), Klucz("Ewa Lis")],
            KeysInData = [Klucz("Jan Kowalski")],
        };

        Assert.True(result.Ok);
        Assert.Equal(new[] { "Ewa Lis" }, result.KeysWithoutData.Select(key => key.Display));
        Assert.Contains("brak danych dla: Ewa Lis", result.Summary());
    }

    [Fact]
    public void Plik_bez_arkusza_z_danymi_przechodzi_i_mowi_o_tym_wprost()
    {
        WorksheetVerificationResult result = Clean() with { DataSheetInFile = false };

        Assert.True(result.Ok);
        Assert.Contains("arkusz z danymi skasowany", result.Summary());
    }

    [Fact]
    public void Brak_arkusza_i_brak_dowodow_to_nie_jest_zgoda_na_wyslanie()
    {
        WorksheetVerificationResult result = Clean() with
        {
            DataSheetInFile = false,
            KeysInData = [],
            KeptRowCount = 0,
            Error = "w pliku nie ma arkusza z danymi, a przy jego tworzeniu nie zebrano dowodow",
        };

        Assert.False(result.Ok);
        Assert.StartsWith("BŁĄD WERYFIKACJI", result.Summary());
    }

    [Fact]
    public void Weryfikacja_bez_dowodow_z_podzialu_odrzuca_plik()
    {
        var config = new SplitConfig { Mode = SplitMode.Worksheet, SplitColumn = "Opiekun" };
        var recipient = new Recipient("Jan Kowalski", [Klucz("Jan Kowalski")], ["Dashboard"]);

        WorksheetVerificationResult result = WorksheetVerifier.VerifyRecipientFile(
            @"C:\out\bez-znaczenia.xlsx", config, recipient,
            [Klucz("Jan Kowalski"), Klucz("Anna Nowak")], evidence: null, NullProgressReporter.Instance);

        Assert.False(result.Ok);
        Assert.Contains("dowod", result.Summary());
    }

    [Fact]
    public void Pisownia_wartosci_odbiorcy_nie_ma_znaczenia()
    {
        WorksheetVerificationResult result = Clean() with { KeysInData = [Klucz("JAN KOWALSKI")] };

        Assert.True(result.Ok);
        Assert.Empty(result.ForeignKeys);
    }

    [Fact]
    public void Obca_kombinacja_zlozona_z_wlasnych_wartosci_odbiorcy_blokuje_plik()
    {
        WorksheetVerificationResult result = Clean() with
        {
            ExpectedKeys = [Klucz("Polnoc", "Kowalski"), Klucz("Poludnie", "Nowak")],
            KeysInData = [Klucz("Polnoc", "Kowalski"), Klucz("Polnoc", "Nowak")],
        };

        Assert.False(result.Ok);
        Assert.Equal(new[] { "Polnoc - Nowak" }, result.ForeignKeys.Select(key => key.Display));
        Assert.Contains("obce kombinacje", result.Summary());
    }

    [Fact]
    public void Pusty_plik_przechodzi_tylko_gdy_zostal_dopuszczony()
    {
        WorksheetVerificationResult empty = Clean() with { KeysInData = [], KeptRowCount = 0 };

        Assert.False(empty.Ok);
        Assert.True((empty with { EmptyAllowed = true }).Ok);
        Assert.Contains("plik pusty", (empty with { EmptyAllowed = true }).Summary());
    }
}
