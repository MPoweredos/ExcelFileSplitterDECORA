using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class SlicerItemsTests
{
    private static HashSet<string> Dane(params string[] values) => new(values, StringComparer.OrdinalIgnoreCase);

    [Fact]
    public void Spacje_na_koncu_podpisu_nie_robia_z_niego_obcej_wartosci()
    {
        var outside = SlicerItems.OutsideData(
            ["PU VOLDEN SOLID ALU 2IN1 1MMX6M CL   "], Dane("PU VOLDEN SOLID ALU 2IN1 1MMX6M CL"));

        Assert.Empty(outside);
    }

    [Fact]
    public void Spacje_na_poczatku_tez_nie()
    {
        Assert.Empty(SlicerItems.OutsideData(["  PH DIY GDAŃSK"], Dane("PH DIY GDAŃSK")));
    }

    [Fact]
    public void Wartosci_spoza_danych_sa_zglaszane()
    {
        var outside = SlicerItems.OutsideData(["PH DIY GDAŃSK", "PH DIY KRAKÓW "], Dane("PH DIY GDAŃSK"));

        Assert.Equal(["PH DIY KRAKÓW"], outside);
    }

    [Fact]
    public void Wielkosc_liter_rozstrzyga_porownanie_zbioru()
    {
        Assert.Empty(SlicerItems.OutsideData(["ph diy gdańsk"], Dane("PH DIY GDAŃSK")));
    }

    [Fact]
    public void Elementy_techniczne_sa_pomijane()
    {
        Assert.Empty(SlicerItems.OutsideData(["(puste)", "(blank)", "   ", ""], Dane("PH DIY GDAŃSK")));
    }

    [Fact]
    public void Dodatkowy_filtr_pomija_wskazane_podpisy()
    {
        var outside = SlicerItems.OutsideData(["2024", "OBI"], Dane("LEROY MERLIN"), caption => caption == "2024");

        Assert.Equal(["OBI"], outside);
    }
}
