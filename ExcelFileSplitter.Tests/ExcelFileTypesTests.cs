using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ExcelFileTypesTests
{
    [Theory]
    [InlineData(@"C:\raporty\Sprzedaz.xlsx")]
    [InlineData(@"C:\raporty\Sprzedaz.xlsm")]
    [InlineData(@"C:\raporty\Sprzedaz.xlsb")]
    [InlineData(@"C:\raporty\Sprzedaz.xls")]
    [InlineData(@"C:\raporty\SPRZEDAZ.XLSX")]
    public void Skoroszyt_Excela_jest_rozpoznawany(string path) =>
        Assert.True(ExcelFileTypes.IsWorkbook(path));

    [Theory]
    [InlineData(@"C:\raporty\Sprzedaz.csv")]
    [InlineData(@"C:\raporty\Sprzedaz.txt")]
    [InlineData(@"C:\raporty\Sprzedaz.pdf")]
    [InlineData(@"C:\raporty\Sprzedaz")]
    [InlineData(@"C:\raporty\Sprzedaz.xlsx.txt")]
    [InlineData("")]
    public void Co_innego_niz_skoroszyt_jest_odrzucane(string path) =>
        Assert.False(ExcelFileTypes.IsWorkbook(path));

    [Fact]
    public void Model_danych_nie_zmiesci_sie_w_starym_xls()
    {
        Assert.True(ExcelFileTypes.IsWorkbook(@"C:\raporty\Raport.xls"));
        Assert.False(ExcelFileTypes.CanHoldDataModel(@"C:\raporty\Raport.xls"));

        Assert.True(ExcelFileTypes.CanHoldDataModel(@"C:\raporty\Raport.xlsm"));
        Assert.True(ExcelFileTypes.CanHoldDataModel(@"C:\raporty\Raport.xlsb"));
    }

    [Fact]
    public void Kontrola_nie_wywraca_sie_na_zadnym_tekscie()
    {
        Assert.False(ExcelFileTypes.IsWorkbook("   "));
        Assert.False(ExcelFileTypes.IsWorkbook("bez rozszerzenia"));
        Assert.False(ExcelFileTypes.IsWorkbook(".xlsx.zip"));

        Assert.True(ExcelFileTypes.IsWorkbook("C:\\raporty\\Raport\0.xlsx"));
    }

    [Fact]
    public void Lista_rozszerzen_do_komunikatu_czyta_sie_po_polsku() =>
        Assert.Equal("xlsx, xlsm, xlsb albo xls", ExcelFileTypes.Listed);
}
