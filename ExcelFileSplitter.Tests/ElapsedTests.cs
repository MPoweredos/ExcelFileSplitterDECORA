using ExcelFileSplitter;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ElapsedTests
{
    [Theory]
    [InlineData(0, "0,0s")]
    [InlineData(0.42, "0,4s")]
    [InlineData(1.5, "1,5s")]
    [InlineData(9.94, "9,9s")]
    public void Krotki_krok_pokazuje_dziesietne_sekundy(double seconds, string expected) =>
        Assert.Equal(expected.Replace(',', Separator), TimeSpan.FromSeconds(seconds).Text());

    [Theory]
    [InlineData(10, "00:10")]
    [InlineData(46, "00:46")]
    [InlineData(86, "01:26")]
    [InlineData(3600 + 125, "02:05")]
    public void Dluzszy_krok_pokazuje_minuty_i_sekundy(double seconds, string expected) =>
        Assert.Equal(expected, TimeSpan.FromSeconds(seconds).Text());

    [Fact]
    public void Ulamek_sekundy_nie_wyglada_na_zero()
    {
        string text = TimeSpan.FromMilliseconds(400).Text();

        Assert.NotEqual("00:00", text);
        Assert.Contains("4", text);
    }

    private static char Separator =>
        System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
}
