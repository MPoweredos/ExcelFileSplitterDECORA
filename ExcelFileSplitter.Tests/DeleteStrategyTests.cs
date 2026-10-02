using ExcelFileSplitter.Interop;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class DeleteStrategyTests
{
    private static DataRange Range(string? tableName = null, int lastColumn = 20) => new(
        SheetName: "Baza",
        TableName: tableName,
        HeaderRow: 1,
        FirstDataRow: 2,
        LastDataRow: 100_000,
        FirstColumn: 1,
        LastColumn: lastColumn,
        Headers: ["PH", "Kwota"]);

    [Fact]
    public void Przy_wielu_blokach_na_zwyklym_zakresie_sortujemy()
    {
        Assert.True(WorksheetTable.CanDeleteBySorting(Range(), entireRow: true, blockCount: 928));
    }

    [Fact]
    public void Przy_kilku_blokach_sortowanie_sie_nie_oplaca()
    {
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(), entireRow: true, blockCount: 3));
    }

    [Fact]
    public void Przy_kilkudziesieciu_blokach_sortowanie_sie_nie_oplaca()
    {
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(), entireRow: true, blockCount: 39));
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(), entireRow: true, blockCount: 102));
    }

    [Fact]
    public void Powyzej_progu_sortowanie_wygrywa()
    {
        Assert.True(WorksheetTable.CanDeleteBySorting(Range(), entireRow: true, blockCount: 530));
    }

    [Fact]
    public void Tabeli_Excela_nie_sortujemy()
    {
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(tableName: "tblSprzedaz"), entireRow: true, blockCount: 928));
    }

    [Fact]
    public void Przy_kasowaniu_samych_komorek_danych_nie_sortujemy()
    {
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(), entireRow: false, blockCount: 928));
    }

    [Fact]
    public void Bez_wolnej_kolumny_za_danymi_nie_sortujemy()
    {
        Assert.False(WorksheetTable.CanDeleteBySorting(Range(lastColumn: 16_384), entireRow: true, blockCount: 928));
    }
}
