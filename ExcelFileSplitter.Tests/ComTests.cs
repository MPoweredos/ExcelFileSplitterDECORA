using System.Runtime.InteropServices;
using ExcelFileSplitter.Interop;
using Microsoft.CSharp.RuntimeBinder;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ComTests
{
    public static TheoryData<Exception> OdmowyExcela =>
    [
        new COMException("Nieokreslony blad"),
        new InvalidCastException(),
        new MissingMemberException(),
        new NotSupportedException(),
        new RuntimeBinderException("nie ma takiej wlasciwosci"),
    ];

    public static TheoryData<Exception> NaszeBledy =>
    [
        new NullReferenceException(),
        new InvalidOperationException("zly stan"),
        new ArgumentNullException("parametr"),
        new IndexOutOfRangeException(),
        new InvalidDataException(),
    ];

    [Theory]
    [MemberData(nameof(OdmowyExcela))]
    public void Odmowa_Excela_jest_rozpoznawana(Exception exception) => Assert.True(Com.SaysNo(exception));

    [Theory]
    [MemberData(nameof(NaszeBledy))]
    public void Nasz_wlasny_blad_nie_jest_odmowa(Exception exception) => Assert.False(Com.SaysNo(exception));

    [Fact]
    public void Try_zwraca_prawde_gdy_krok_sie_wykonal()
    {
        bool wykonane = false;

        Assert.True(Com.Try(() => wykonane = true));
        Assert.True(wykonane);
    }

    [Theory]
    [MemberData(nameof(OdmowyExcela))]
    public void Try_zwraca_falsz_gdy_Excel_odmowil(Exception exception) =>
        Assert.False(Com.Try(() => throw exception));

    [Theory]
    [MemberData(nameof(NaszeBledy))]
    public void Try_przepuszcza_nasz_wlasny_blad(Exception exception) =>
        Assert.Same(exception, Assert.ThrowsAny<Exception>(() => Com.Try(() => throw exception)));

    [Fact]
    public void Or_zwraca_odczytana_wartosc()
    {
        Assert.Equal(42, Com.Or(() => 42, 0));
    }

    [Theory]
    [MemberData(nameof(OdmowyExcela))]
    public void Or_zwraca_wartosc_zastepcza_gdy_Excel_odmowil(Exception exception) =>
        Assert.Equal(-1, Com.Or<int>(() => throw exception, -1));

    [Theory]
    [MemberData(nameof(NaszeBledy))]
    public void Or_przepuszcza_nasz_wlasny_blad(Exception exception) =>
        Assert.Same(exception, Assert.ThrowsAny<Exception>(() => Com.Or<int>(() => throw exception, -1)));

    public static TheoryData<Exception> SmiercExcela =>
    [
        new COMException("Serwer RPC jest niedostepny", unchecked((int)0x800706BA)),
        new COMException("Wywolanie zdalnej procedury nie powiodlo sie", unchecked((int)0x800706BE)),
        new COMException("Wywolanie zdalnej procedury nie powiodlo sie i nie zostalo wykonane", unchecked((int)0x800706BF)),
        new COMException("Obiekt odlaczyl sie od klientow", unchecked((int)0x80010108)),
        new COMException("Serwer wywrocil sie w trakcie wywolania", unchecked((int)0x80010105)),
        new COMException("Obiekt nie jest juz podlaczony", unchecked((int)0x800401FD)),
        new InvalidComObjectException(),
    ];

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Smierc_procesu_Excela_jest_rozpoznawana(Exception exception) =>
        Assert.True(Com.SessionIsDead(exception));

    [Theory]
    [MemberData(nameof(OdmowyExcela))]
    public void Zwykla_odmowa_nie_jest_smiercia_sesji(Exception exception) =>
        Assert.False(Com.SessionIsDead(exception));

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Try_przepuszcza_smierc_sesji(Exception exception)
    {
        Assert.Same(exception, Assert.ThrowsAny<Exception>(() => Com.Try(() => throw exception)));
    }

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Or_przepuszcza_smierc_sesji(Exception exception) =>
        Assert.Same(exception, Assert.ThrowsAny<Exception>(() => Com.Or<int>(() => throw exception, -1)));

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Sprzatanie_przelyka_takze_smierc_sesji(Exception exception) =>
        Assert.False(Com.WhenClosing(() => throw exception));

    [Theory]
    [MemberData(nameof(NaszeBledy))]
    public void Sprzatanie_przepuszcza_nasz_wlasny_blad(Exception exception) =>
        Assert.Same(exception, Assert.ThrowsAny<Exception>(() => Com.WhenClosing(() => throw exception)));

    [Fact]
    public void Sprzatanie_zwraca_prawde_gdy_sie_udalo()
    {
        bool posprzatane = false;

        Assert.True(Com.WhenClosing(() => posprzatane = true));
        Assert.True(posprzatane);
    }

    [Theory]
    [MemberData(nameof(OdmowyExcela))]
    public void Odmowa_Excela_zasluguje_na_powtorke_na_swiezym_procesie(Exception exception)
    {
        Assert.True(Com.WorthRetryingOnFreshExcel(exception));
    }

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Smierc_Excela_zasluguje_na_powtorke_na_swiezym_procesie(Exception exception) =>
        Assert.True(Com.WorthRetryingOnFreshExcel(exception));

    [Theory]
    [MemberData(nameof(NaszeBledy))]
    public void Nasz_wlasny_blad_nie_zasluguje_na_restart_Excela(Exception exception)
    {
        Assert.False(Com.WorthRetryingOnFreshExcel(exception));
    }

    [Fact]
    public void Anulowanie_i_zamknieta_sesja_nie_zasluguja_na_restart_Excela()
    {
        Assert.False(Com.WorthRetryingOnFreshExcel(new OperationCanceledException()));
        Assert.False(Com.WorthRetryingOnFreshExcel(new ObjectDisposedException("SplitWorkspace")));
    }

    [Theory]
    [MemberData(nameof(SmiercExcela))]
    public void Powtorke_widac_takze_przez_opakowanie(Exception exception)
    {
        var opakowany = new InvalidOperationException("Excel przestal odpowiadac", exception);

        Assert.True(Com.WorthRetryingOnFreshExcel(opakowany));
    }
}
