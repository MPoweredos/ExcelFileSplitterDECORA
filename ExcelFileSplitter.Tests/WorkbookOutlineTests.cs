using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class WorkbookOutlineTests
{
    private static DataCandidate Tabela(string name, string sheet, long rows) =>
        new(name, sheet, HeaderRow: 1, rows, ["PH", "Kwota"]);

    private static DataCandidate Arkusz(string sheet, long rows, int headerRow = 1) =>
        new(null, sheet, headerRow, rows, ["PH", "Kwota"]);

    private static SplitKey Klucz(params string[] values) => SplitKey.Of(values);

    [Fact]
    public void Zliczanie_laczy_kombinacje_rozniace_sie_wielkoscia_liter()
    {
        List<SplitKeyCount> keys = SplitKeyCount.Tally(
            [Klucz("PH1"), Klucz("ph1"), Klucz("PH2"), Klucz("Ph1")]);

        Assert.Equal(2, keys.Count);
        Assert.Equal("PH1", keys[0].Display);
        Assert.Equal<long?>(3L, keys[0].RowCount);
        Assert.Equal("PH2", keys[1].Display);
        Assert.Equal<long?>(1L, keys[1].RowCount);
    }

    [Fact]
    public void Zliczanie_pokazuje_pisownie_z_pierwszego_wystapienia()
    {
        List<SplitKeyCount> keys = SplitKeyCount.Tally([Klucz("ph duzy"), Klucz("PH DUZY")]);

        Assert.Equal("ph duzy", Assert.Single(keys).Display);
    }

    [Fact]
    public void Zliczanie_pomija_wiersze_z_niekompletnym_kluczem()
    {
        List<SplitKeyCount> keys = SplitKeyCount.Tally(
            [Klucz("Polnoc", "Kowalski"), Klucz("Polnoc", ""), Klucz("", "Nowak"), Klucz("Polnoc", "Kowalski")]);

        Assert.Equal<long?>(2L, Assert.Single(keys).RowCount);
    }

    [Fact]
    public void Zliczanie_rozroznia_kombinacje_zlozone_z_tych_samych_wartosci()
    {
        List<SplitKeyCount> keys = SplitKeyCount.Tally(
            [Klucz("Polnoc", "Kowalski"), Klucz("Poludnie", "Nowak"), Klucz("Polnoc", "Nowak")]);

        Assert.Equal(3, keys.Count);
    }

    [Fact]
    public void Zliczanie_ustawia_kombinacje_alfabetycznie()
    {
        List<SplitKeyCount> keys = SplitKeyCount.Tally([Klucz("zeta"), Klucz("Alfa"), Klucz("beta")]);

        Assert.Equal(new[] { "Alfa", "beta", "zeta" }, keys.Select(key => key.Display));
    }

    [Fact]
    public void Podpowiedz_wybiera_tabele_Excela_przed_surowym_arkuszem()
    {
        DataCandidate? best = WorkbookOutline.Best([Arkusz("Baza", 500_000), Tabela("tblSprzedaz", "Dane", 1_000)]);

        Assert.Equal("tblSprzedaz", best!.TableName);
    }

    [Fact]
    public void Podpowiedz_wybiera_najwiekszy_zakres_gdy_nie_ma_tabel()
    {
        DataCandidate? best = WorkbookOutline.Best([Arkusz("Slownik", 40), Arkusz("Baza", 120_000)]);

        Assert.Equal("Baza", best!.SheetName);
    }

    [Fact]
    public void Podpowiedz_jest_pusta_gdy_nie_ma_zadnych_kandydatow()
    {
        Assert.Null(WorkbookOutline.Best([]));
    }

    [Fact]
    public void Tabela_trafia_do_konfiguracji_jako_Table_bez_arkusza()
    {
        WorksheetOptions source = Tabela("tblSprzedaz", "Dane", 10).ToWorksheetOptions();

        Assert.Equal("tblSprzedaz", source.Table);
        Assert.Equal("", source.Sheet);
        Assert.True(source.DeleteEntireRow);
    }

    [Fact]
    public void Surowy_zakres_trafia_do_konfiguracji_jako_arkusz_z_wierszem_naglowkow()
    {
        WorksheetOptions source = Arkusz("Baza", 10, headerRow: 3).ToWorksheetOptions(deleteEntireRow: false);

        Assert.Equal("", source.Table);
        Assert.Equal("Baza", source.Sheet);
        Assert.Equal(3, source.HeaderRow);
        Assert.False(source.DeleteEntireRow);
    }
}
