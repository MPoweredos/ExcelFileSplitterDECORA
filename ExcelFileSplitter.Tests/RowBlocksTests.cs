using Xunit;

namespace ExcelFileSplitter.Tests;

public class RowBlocksTests
{
    private static Func<string, bool> Keep(params string[] values)
    {
        var kept = new HashSet<string>(values, StringComparer.OrdinalIgnoreCase);
        return kept.Contains;
    }

    [Fact]
    public void Brak_wierszy_to_brak_blokow()
    {
        Assert.Empty(RowBlocks.ToDelete([], 2, Keep("Jan")));
    }

    [Fact]
    public void Gdy_wszystko_zostaje_nie_ma_czego_kasowac()
    {
        var blocks = RowBlocks.ToDelete(["Jan", "Jan", "Jan"], 2, Keep("Jan"));

        Assert.Empty(blocks);
    }

    [Fact]
    public void Gdy_wszystko_znika_powstaje_jeden_blok()
    {
        var blocks = RowBlocks.ToDelete(["Anna", "Piotr", "Ewa"], 2, Keep("Jan"));

        RowBlock block = Assert.Single(blocks);
        Assert.Equal(2, block.First);
        Assert.Equal(4, block.Last);
        Assert.Equal(3, block.Count);
    }

    [Fact]
    public void Numeracja_idzie_za_pierwszym_wierszem_danych()
    {
        var blocks = RowBlocks.ToDelete(["Anna", "Jan"], 8, Keep("Jan"));

        RowBlock block = Assert.Single(blocks);
        Assert.Equal(8, block.First);
        Assert.Equal(8, block.Last);
    }

    [Fact]
    public void Blok_na_poczatku_i_na_koncu_jest_wykrywany()
    {
        var blocks = RowBlocks.ToDelete(["Anna", "Jan", "Jan", "Ewa", "Piotr"], 2, Keep("Jan"));

        Assert.Equal(2, blocks.Count);
        Assert.Equal(new RowBlock(2, 2), blocks[0]);
        Assert.Equal(new RowBlock(5, 6), blocks[1]);
    }

    [Fact]
    public void Wiersze_na_przemian_daja_blok_na_kazdy_wiersz()
    {
        var blocks = RowBlocks.ToDelete(["Jan", "Anna", "Jan", "Anna", "Jan"], 2, Keep("Jan"));

        Assert.Equal(2, blocks.Count);
        Assert.Equal(new RowBlock(3, 3), blocks[0]);
        Assert.Equal(new RowBlock(5, 5), blocks[1]);
    }

    [Fact]
    public void Liczba_blokow_zalezy_od_wierszy_zostajacych_a_nie_kasowanych()
    {
        var values = new List<string>();
        for (int i = 0; i < 1000; i++) values.Add(i is 400 or 700 ? "Jan" : "Anna");

        var blocks = RowBlocks.ToDelete(values, 2, Keep("Jan"));

        Assert.Equal(3, blocks.Count);
        Assert.Equal(998, RowBlocks.TotalRows(blocks));
    }

    [Fact]
    public void Pisownia_wartosci_nie_ma_znaczenia()
    {
        var blocks = RowBlocks.ToDelete(["JAN KOWALSKI", "jan kowalski", "Anna"], 2, Keep("Jan Kowalski"));

        RowBlock block = Assert.Single(blocks);
        Assert.Equal(new RowBlock(4, 4), block);
    }

    [Fact]
    public void Puste_komorki_sa_kasowane_jak_kazda_obca_wartosc()
    {
        var blocks = RowBlocks.ToDelete(["Jan", "", "Jan"], 2, Keep("Jan"));

        RowBlock block = Assert.Single(blocks);
        Assert.Equal(new RowBlock(3, 3), block);
    }

    [Fact]
    public void Plik_zbiorczy_zostawia_wiersze_wszystkich_swoich_wartosci()
    {
        var blocks = RowBlocks.ToDelete(["Jan", "Anna", "Ewa", "Jan"], 2, Keep("Jan", "Ewa"));

        RowBlock block = Assert.Single(blocks);
        Assert.Equal(new RowBlock(3, 3), block);
    }
}
