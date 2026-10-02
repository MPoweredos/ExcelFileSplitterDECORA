using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class PivotRefreshTests
{
    private static readonly DateTime Start = new(2026, 9, 30, 10, 0, 0);

    private static PivotRefresh Refresh(DateTime started, params DateTime[] failedDates) =>
        new(started, Refreshed: 1,
            Failures: failedDates.Select(_ => new PivotRefreshFailure([], "błąd", null)).ToList(),
            FailedDates: failedDates.ToHashSet());

    private static PivotRefreshFailure Failure(params PivotTableRef[] tables) =>
        new(tables, "This command requires at least two rows of source data", "Dane!A1:O1 (sam wiersz nagłówków, bez danych)");

    private static HashSet<string> Arkusze(params string[] names) => new(names, StringComparer.OrdinalIgnoreCase);

    [Fact]
    public void Cache_odswiezony_w_trakcie_przebiegu_jest_swiezy()
    {
        Assert.False(Refresh(Start).IsStale(Start.AddSeconds(3)));
    }

    [Fact]
    public void Cache_sprzed_przebiegu_jest_stary()
    {
        Assert.True(Refresh(Start).IsStale(Start.AddMinutes(-10)));
    }

    [Fact]
    public void Data_z_tej_samej_sekundy_co_start_nie_jest_alarmem()
    {
        Assert.False(Refresh(Start.AddMilliseconds(750)).IsStale(Start));
    }

    [Fact]
    public void Cache_ktory_sie_nie_odswiezyl_jest_stary_nawet_z_data_z_przyszlosci()
    {
        DateTime future = Start.AddHours(3);

        Assert.True(Refresh(Start, future).IsStale(future));
    }

    [Fact]
    public void Data_nieudanego_cache_zapisana_z_mniejsza_dokladnoscia_nadal_go_rozpoznaje()
    {
        DateTime future = Start.AddHours(3);

        Assert.True(Refresh(Start, future).IsStale(future.AddMilliseconds(4)));
    }

    [Fact]
    public void Swiezy_cache_nie_jest_mylony_z_nieudanym_z_sasiedniej_sekundy()
    {
        Assert.False(Refresh(Start, Start.AddSeconds(2)).IsStale(Start.AddSeconds(3)));
    }

    [Fact]
    public void Brak_daty_przy_nieudanych_odswiezeniach_blokuje()
    {
        Assert.True(Refresh(Start, Start.AddDays(-2)).IsStale(null));
    }

    [Fact]
    public void Brak_daty_gdy_wszystko_sie_odswiezylo_nie_blokuje()
    {
        Assert.False(Refresh(Start).IsStale(null));
    }

    [Fact]
    public void Tabela_na_skasowanym_arkuszu_to_informacja_a_nie_ostrzezenie()
    {
        (bool isWarning, string message) =
            Failure(new PivotTableRef("Tabela przestawna1", "Arkusz1")).Report(Arkusze("Raport"));

        Assert.False(isWarning);
        Assert.Contains("'Tabela przestawna1' (arkusz 'Arkusz1')", message);
        Assert.Contains("nie trafia do odbiorcy", message);
        Assert.Contains("at least two rows", message);
        Assert.Contains("źródło danych: Dane!A1:O1", message);
    }

    [Fact]
    public void Tabela_na_arkuszu_ktory_zostaje_to_ostrzezenie()
    {
        (bool isWarning, string message) =
            Failure(new PivotTableRef("Tabela przestawna1", "raport")).Report(Arkusze("Raport"));

        Assert.True(isWarning);
        Assert.Contains("zostaje w pliku", message);
    }

    [Fact]
    public void Ostrzezenie_wymienia_tylko_tabele_ktore_zostaja()
    {
        PivotRefreshFailure failure = Failure(
            new PivotTableRef("Tabela przestawna1", "Arkusz1"), new PivotTableRef("Tabela przestawna2", "Raport"));

        (bool isWarning, string message) = failure.Report(Arkusze("Raport"));

        Assert.True(isWarning);
        Assert.Contains("Tabela przestawna2", message);
        Assert.DoesNotContain("Tabela przestawna1", message);
    }

    [Fact]
    public void Kilka_tabel_na_jednym_skasowanym_arkuszu()
    {
        PivotRefreshFailure failure = Failure(
            new PivotTableRef("Tabela przestawna1", "Arkusz1"), new PivotTableRef("Tabela przestawna2", "Arkusz1"));

        (bool isWarning, string message) = failure.Report(Arkusze("Raport"));

        Assert.False(isWarning);
        Assert.Contains("ich arkusz nie trafia do odbiorcy", message);
    }

    [Fact]
    public void Cache_bez_tabel_to_ostrzezenie()
    {
        (bool isWarning, string message) = Failure().Report(Arkusze("Raport"));

        Assert.True(isWarning);
        Assert.Contains("nie da się przypisać żadnej tabeli", message);
    }

    [Theory]
    [InlineData("Dane!R1C1:R500C15", "Dane!A1:O500")]
    [InlineData("Dane!R1C1:R1C15", "Dane!A1:O1 (sam wiersz nagłówków, bez danych)")]
    [InlineData("'Moje dane'!R2C27:R10C28", "'Moje dane'!AA2:AB10")]
    [InlineData("Dane!R1C1", "Dane!A1 (sam wiersz nagłówków, bez danych)")]
    [InlineData("Tabela1", "Tabela1")]
    [InlineData(" DaneRaportu ", "DaneRaportu")]
    public void Zrodlo_danych_jest_pokazywane_jak_w_Excelu(string sourceData, string expected)
    {
        Assert.Equal(expected, PivotSource.Describe(sourceData));
    }

    [Fact]
    public void Zrodlo_ktore_nie_jest_tekstem_jest_pomijane()
    {
        Assert.Null(PivotSource.Describe(null));
        Assert.Null(PivotSource.Describe(new object[] { "ODBC;DSN=Hurtownia", "SELECT 1" }));
        Assert.Null(PivotSource.Describe("  "));
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16_384, "XFD")]
    public void Numer_kolumny_na_litery(int column, string expected)
    {
        Assert.Equal(expected, PivotSource.ColumnLetters(column));
    }

    [Fact]
    public void Z_dlugiego_komunikatu_Excela_zostaje_pierwsze_zdanie()
    {
        const string message =
            "This command requires at least two rows of source data. You cannot use the command on a selection " +
            "in only one row. Try the following:\n\n• If you're using an advanced filter, select a range of cells " +
            "that contains at least two rows of data.";

        Assert.Equal("This command requires at least two rows of source data", ExcelMessage.FirstSentence(message));
    }

    [Theory]
    [InlineData("Reference isn't valid.", "Reference isn't valid")]
    [InlineData("Exception from HRESULT: 0x800A03EC", "Exception from HRESULT: 0x800A03EC")]
    [InlineData(@"Cannot open PivotTable source file 'C:\raporty\plik.xlsx'.", @"Cannot open PivotTable source file 'C:\raporty\plik.xlsx'")]
    [InlineData("  ", "bez opisu błędu")]
    public void Krotki_komunikat_zostaje_bez_kropki_na_koncu(string message, string expected)
    {
        Assert.Equal(expected, ExcelMessage.FirstSentence(message));
    }
}
