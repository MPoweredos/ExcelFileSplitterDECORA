using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class StaleOutputTests : IDisposable
{
    private readonly string _folder;

    public StaleOutputTests()
    {
        _folder = Path.Combine(Path.GetTempPath(), $"efs-stale-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_folder);
    }

    private PassRun Run(params string[] recipients)
    {
        var config = new SplitConfig
        {
            Mode = SplitMode.Worksheet,
            SourceWorkbook = @"C:\Raporty\Sprzedaz.xlsb",
            OutputFolder = _folder,
            FileNameTemplate = "Raport_{value}.xlsx",
        };

        List<Recipient> files = recipients
            .Select(name => new Recipient(name, [SplitKey.Of(name)], ["Dashboard"]))
            .ToList();

        var info = new WorkbookInfo([], [], [], [], [], new Dictionary<string, string>(), [], [], [], null);
        return new PassRun(new SplitPass(), config, info, new RecipientPlan(files, [], []));
    }

    private void Put(string fileName) => File.WriteAllText(Path.Combine(_folder, fileName), "");

    [Fact]
    public void Plik_z_poprzedniego_przebiegu_jest_zglaszany()
    {
        Put("Raport_ANNA.xlsb");
        Put("Raport_STARY MENADZER.xlsb");

        List<StaleOutput> stale = SplitRun.StaleOutputFiles([Run("ANNA")]);

        Assert.Single(stale);
        Assert.Equal("Raport_STARY MENADZER.xlsb", Assert.Single(stale[0].Files));
        Assert.Contains("NIE powstaną", stale[0].Describe());
    }

    [Fact]
    public void Plik_ktory_wlasnie_powstanie_nie_jest_stary()
    {
        Put("Raport_ANNA.xlsb");
        Put("Raport_PAWEL.xlsb");

        Assert.Empty(SplitRun.StaleOutputFiles([Run("ANNA", "PAWEL")]));
    }

    [Fact]
    public void Plik_odrzucony_przez_weryfikacje_nie_jest_stary()
    {
        Put("Raport_ANNA_NIE_WYSYŁAĆ.xlsb");

        Assert.Empty(SplitRun.StaleOutputFiles([Run("ANNA")]));
    }

    [Fact]
    public void Plik_odrzucony_ze_starym_sufiksem_jest_zglaszany()
    {
        Put("Raport_ANNA_NIE_WYSYLAC.xlsb");

        Assert.Equal("Raport_ANNA_NIE_WYSYLAC.xlsb", Assert.Single(Assert.Single(SplitRun.StaleOutputFiles([Run("ANNA")])).Files));
    }

    [Fact]
    public void Nie_excelowe_pliki_i_blokady_Excela_sa_pomijane()
    {
        Put("notatki.txt");
        Put("~$Raport_ANNA.xlsb");

        Assert.Empty(SplitRun.StaleOutputFiles([Run("ANNA")]));
    }

    [Fact]
    public void Brak_folderu_wynikowego_nie_jest_bledem()
    {
        Directory.Delete(_folder, recursive: true);

        Assert.Empty(SplitRun.StaleOutputFiles([Run("ANNA")]));
    }

    public void Dispose()
    {
        try { Directory.Delete(_folder, recursive: true); } catch { }
    }
}
