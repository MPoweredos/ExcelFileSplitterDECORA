using ExcelFileSplitter.Interop;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ExcelProcessRegistryTests
{
    private const long ExcelStart = 638_000_000_000_000_000;
    private const long OwnerStart = 637_999_999_000_000_000;

    private static Func<int, RunningProcess?> Processes(params (int Pid, string Name, long Start)[] running) =>
        pid => running.Where(process => process.Pid == pid)
            .Select(process => new RunningProcess(process.Name, process.Start))
            .FirstOrDefault();

    [Fact]
    public void Wpis_zapisany_i_odczytany_jest_taki_sam()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);

        Assert.Equal(entry, ExcelProcessEntry.Parse(entry.Format()));
    }

    [Theory]
    [InlineData("")]
    [InlineData("4321;638000000000000000;1234")]
    [InlineData("4321;abc;1234;637999999000000000")]
    [InlineData("0;638000000000000000;1234;637999999000000000")]
    [InlineData("4321;638000000000000000;-5;637999999000000000")]
    public void Uszkodzony_wpis_jest_pomijany(string line)
    {
        Assert.Null(ExcelProcessEntry.Parse(line));
    }

    [Fact]
    public void Excel_po_programie_ktory_juz_nie_dziala_jest_do_zamkniecia()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);
        var lookup = Processes((4321, "EXCEL", ExcelStart));

        var (kill, keep) = ExcelProcessRegistry.Decide([entry], lookup, e => ExcelProcessRegistry.OwnerIsGone(e, lookup));

        Assert.Equal([entry], kill);
        Assert.Empty(keep);
    }

    [Fact]
    public void Excel_dzialajacego_programu_zostaje()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);
        var lookup = Processes((4321, "EXCEL", ExcelStart), (1234, "ExcelFileSplitter.Ui", OwnerStart));

        var (kill, keep) = ExcelProcessRegistry.Decide([entry], lookup, e => ExcelProcessRegistry.OwnerIsGone(e, lookup));

        Assert.Empty(kill);
        Assert.Equal([entry], keep);
    }

    [Fact]
    public void Numer_programu_przejety_przez_inny_proces_znaczy_ze_program_juz_nie_dziala()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);
        var lookup = Processes((4321, "EXCEL", ExcelStart), (1234, "chrome", OwnerStart + 99));

        Assert.True(ExcelProcessRegistry.OwnerIsGone(entry, lookup));
    }

    [Theory]
    [InlineData("EXCEL", ExcelStart + 1)]
    [InlineData("notepad", ExcelStart)]
    public void Numer_Excela_przejety_przez_inny_proces_nie_jest_ruszany(string name, long start)
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);
        var lookup = Processes((4321, name, start));

        var (kill, keep) = ExcelProcessRegistry.Decide([entry], lookup, _ => true);

        Assert.Empty(kill);
        Assert.Empty(keep);
    }

    [Fact]
    public void Excel_ktorego_juz_nie_ma_znika_z_listy()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);

        var (kill, keep) = ExcelProcessRegistry.Decide([entry], Processes(), _ => true);

        Assert.Empty(kill);
        Assert.Empty(keep);
    }

    [Fact]
    public void Wielkosc_liter_w_nazwie_procesu_Excela_nie_ma_znaczenia()
    {
        var entry = new ExcelProcessEntry(4321, ExcelStart, 1234, OwnerStart);

        var (kill, _) = ExcelProcessRegistry.Decide([entry, entry], Processes((4321, "excel", ExcelStart)), _ => true);

        Assert.Equal([entry], kill);
    }
}
