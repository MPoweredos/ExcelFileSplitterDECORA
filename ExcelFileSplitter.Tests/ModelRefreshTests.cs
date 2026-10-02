using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ModelRefreshTests
{
    private static readonly DateTime Start = new(2026, 9, 21, 10, 0, 0);

    private static List<StaleTable> Check(DateTime started, params (string Table, DateTime? Refreshed)[] afterRefresh) =>
        ModelSchema.FindTablesNotRefreshed(
            started,
            afterRefresh.ToDictionary(entry => entry.Table, entry => entry.Refreshed, StringComparer.OrdinalIgnoreCase),
            afterRefresh.Select(entry => entry.Table));

    [Fact]
    public void Tabela_z_data_sprzed_przebiegu_jest_zglaszana()
    {
        var stale = Check(Start, ("fact_Zamowienia", Start.AddMinutes(-10)));

        Assert.Single(stale);
        Assert.Equal("fact_Zamowienia", stale[0].Name);
    }

    [Fact]
    public void Tabela_odswiezona_po_starcie_jest_czysta()
    {
        Assert.Empty(Check(Start, ("fact_Zamowienia", Start.AddSeconds(5))));
    }

    [Fact]
    public void Data_z_tej_samej_sekundy_co_start_nie_jest_alarmem()
    {
        var started = Start.AddMilliseconds(750);

        Assert.Empty(Check(started, ("fact_Zamowienia", Start)));
    }

    [Fact]
    public void Sekunda_zapasu_nie_siega_dalej()
    {
        var stale = Check(Start, ("fact_Zamowienia", Start.AddSeconds(-2)));

        Assert.Single(stale);
    }

    [Fact]
    public void Brak_daty_nie_jest_zglaszany()
    {
        Assert.Empty(Check(Start, ("fact_Zamowienia", null)));
    }

    [Fact]
    public void Tabela_nieobecna_w_odczycie_nie_jest_zglaszana()
    {
        var afterRefresh = new Dictionary<string, DateTime?> { ["dim_Klienci"] = Start.AddSeconds(5) };

        Assert.Empty(ModelSchema.FindTablesNotRefreshed(Start, afterRefresh, ["fact_Zamowienia"]));
    }

    [Fact]
    public void Sprawdzamy_tylko_tabele_z_listy()
    {
        var afterRefresh = new Dictionary<string, DateTime?>
        {
            ["fact_Zamowienia"] = Start.AddSeconds(5),
            ["dim_Slownik"] = Start.AddDays(-30),
        };

        Assert.Empty(ModelSchema.FindTablesNotRefreshed(Start, afterRefresh, ["fact_Zamowienia"]));
    }

    [Fact]
    public void Data_z_przyszlosci_przechodzi_i_tak_ma_byc()
    {
        Assert.Empty(Check(Start, ("fact_Zamowienia", Start.AddHours(3))));
    }

    [Fact]
    public void Zgloszenie_niesie_date_tabeli()
    {
        var stale = Check(Start, ("fact_Zamowienia", new DateTime(2026, 9, 20, 8, 15, 30)));

        Assert.Contains("fact_Zamowienia", stale[0].ToString());
        Assert.Contains("2026-09-20 08:15:30", stale[0].ToString());
    }
}
