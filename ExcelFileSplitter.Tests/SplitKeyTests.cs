using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class SplitKeyTests
{
    [Fact]
    public void Wielkosc_liter_nie_rozroznia_kluczy()
    {
        Assert.Equal(SplitKey.Of("polnoc", "kowalski"), SplitKey.Of("POLNOC", "Kowalski"));
    }

    [Fact]
    public void Spacje_na_brzegach_nie_rozrozniaja_kluczy()
    {
        Assert.Equal(SplitKey.Of("Polnoc"), SplitKey.Of("  Polnoc "));
    }

    [Fact]
    public void Kolejnosc_wartosci_ma_znaczenie()
    {
        Assert.NotEqual(SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("Kowalski", "Polnoc"));
    }

    [Fact]
    public void Klucz_o_innej_dlugosci_nie_jest_rowny()
    {
        Assert.NotEqual(SplitKey.Of("Polnoc"), SplitKey.Of("Polnoc", "Kowalski"));
    }

    [Fact]
    public void Rowne_klucze_maja_ten_sam_skrot()
    {
        Assert.Equal(SplitKey.Of("POLNOC", "kowalski").GetHashCode(), SplitKey.Of("polnoc", "KOWALSKI").GetHashCode());
    }

    [Fact]
    public void Zbior_traktuje_rownowazne_klucze_jak_jeden()
    {
        var keys = new HashSet<SplitKey> { SplitKey.Of("Polnoc", "Kowalski"), SplitKey.Of("POLNOC", "KOWALSKI") };

        Assert.Single(keys);
    }

    [Fact]
    public void Klucz_z_pusta_wartoscia_jest_niekompletny()
    {
        Assert.True(SplitKey.Of("Polnoc", "").IsIncomplete);
        Assert.True(SplitKey.Of("", "Kowalski").IsIncomplete);
        Assert.False(SplitKey.Of("Polnoc", "Kowalski").IsIncomplete);
    }

    [Fact]
    public void Opis_sklada_wartosci_w_kolejnosci()
    {
        Assert.Equal("Polnoc - Kowalski", SplitKey.Of("Polnoc", "Kowalski").Display);
    }

    [Fact]
    public void Keys_z_konfiguracji_wygrywa_nad_Values()
    {
        List<SplitKey> keys = SplitKey.From([["Polnoc", "Kowalski"]], ["cokolwiek"]);

        Assert.Equal(SplitKey.Of("Polnoc", "Kowalski"), Assert.Single(keys));
    }

    [Fact]
    public void Bez_Keys_uzywane_sa_Values_jako_klucze_jednoelementowe()
    {
        List<SplitKey> keys = SplitKey.From([], ["Kowalski", "Nowak"]);

        Assert.Equal(new[] { "Kowalski", "Nowak" }, keys.Select(key => key.Display));
        Assert.All(keys, key => Assert.Equal(1, key.Count));
    }
}
