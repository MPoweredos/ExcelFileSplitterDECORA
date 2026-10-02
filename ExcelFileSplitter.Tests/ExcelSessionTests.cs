using ExcelFileSplitter.Interop;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ExcelSessionTests
{
    [Theory]
    [InlineData(0, false)]
    [InlineData(-1, false)]
    [InlineData(1, true)]
    [InlineData(255, true)]
    [InlineData(unchecked((int)0xC0000005), true)]
    [InlineData(unchecked((int)0xC0000409), true)]
    [InlineData(unchecked((int)0x80000003), true)]
    public void Awaria_rozpoznawana_po_kodzie_wyjscia_Excela(int exitCode, bool crash)
    {
        Assert.Equal(crash, ExcelSession.IsCrash(exitCode));
    }
}
