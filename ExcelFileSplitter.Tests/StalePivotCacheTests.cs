using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class StalePivotCacheTests : IDisposable
{
    private static readonly DateTime Start = new(2026, 9, 30, 10, 0, 0);

    private readonly List<string> _createdFiles = [];

    private string Package(params (string Part, string Xml)[] parts)
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.zip");
        _createdFiles.Add(path);

        using var package = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach ((string part, string xml) in parts)
        {
            using Stream stream = package.CreateEntry(part).Open();
            byte[] content = Encoding.UTF8.GetBytes(xml);
            stream.Write(content, 0, content.Length);
        }
        return path;
    }

    private static string Definition(DateTime? refreshedAt, string sourceType = "worksheet") =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" " +
        "refreshedBy=\"Michał\" " +
        (refreshedAt is null ? "" : $"refreshedDate=\"{refreshedAt.Value.ToOADate().ToString("R", CultureInfo.InvariantCulture)}\" ") +
        "createdVersion=\"8\" refreshedVersion=\"8\" minRefreshableVersion=\"3\" recordCount=\"2\">" +
        $"<cacheSource type=\"{sourceType}\"><worksheetSource ref=\"A1:C3\" sheet=\"Dane\"/></cacheSource>" +
        "<cacheFields count=\"1\"><cacheField name=\"Opiekun\" numFmtId=\"0\"><sharedItems/></cacheField></cacheFields>" +
        "</pivotCacheDefinition>";

    private static PivotRefresh Refresh(params DateTime[] failedDates) =>
        new(Start, Refreshed: 1,
            Failures: failedDates.Select(_ => new PivotRefreshFailure([], "błąd", null)).ToList(),
            FailedDates: failedDates.ToHashSet());

    private const string Part = "xl/pivotCache/pivotCacheDefinition1.xml";

    [Fact]
    public void Cache_odswiezony_w_trakcie_podzialu_przechodzi()
    {
        string path = Package((Part, Definition(Start.AddSeconds(5))));

        Assert.Empty(PackageScanner.FindStalePivotCaches(path, Refresh().IsStale));
    }

    [Fact]
    public void Cache_sprzed_podzialu_w_zapisanym_pliku_jest_zgloszony()
    {
        string path = Package((Part, Definition(new DateTime(2026, 9, 29, 14, 2, 11))));

        List<string> stale = PackageScanner.FindStalePivotCaches(path, Refresh().IsStale);

        string entry = Assert.Single(stale);
        Assert.Contains("pivotCacheDefinition1.xml", entry);
        Assert.Contains("2026-09-29 14:02:11", entry);
    }

    [Fact]
    public void Cache_ktory_sie_nie_odswiezyl_jest_zgloszony_nawet_z_data_z_przyszlosci()
    {
        DateTime future = Start.AddHours(3);
        string path = Package((Part, Definition(future)));

        Assert.Single(PackageScanner.FindStalePivotCaches(path, Refresh(future).IsStale));
    }

    [Fact]
    public void Cache_z_zewnetrznego_zrodla_jest_pomijany()
    {
        string path = Package((Part, Definition(Start.AddDays(-30), sourceType: "external")));

        Assert.Empty(PackageScanner.FindStalePivotCaches(path, Refresh().IsStale));
    }

    [Fact]
    public void Brak_daty_blokuje_tylko_gdy_cos_sie_nie_odswiezylo()
    {
        string path = Package((Part, Definition(refreshedAt: null)));

        Assert.Empty(PackageScanner.FindStalePivotCaches(path, Refresh().IsStale));
        Assert.Contains("brak daty", Assert.Single(PackageScanner.FindStalePivotCaches(path, Refresh(Start.AddDays(-1)).IsStale)));
    }

    [Fact]
    public void Uszkodzona_definicja_cache_blokuje_plik()
    {
        string path = Package((Part, "<pivotCacheDefinition refreshedDate=\"1\"><cacheSource"));

        Assert.Contains("nie da się odczytać", Assert.Single(PackageScanner.FindStalePivotCaches(path, Refresh().IsStale)));
    }

    [Fact]
    public void Inne_czesci_paczki_sa_pomijane()
    {
        string path = Package(
            ("xl/worksheets/sheet1.xml", Definition(Start.AddDays(-30))),
            ("xl/pivotCache/pivotCacheRecords1.xml", "<pivotCacheRecords/>"),
            ("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "<Relationships/>"),
            (Part, Definition(Start.AddSeconds(2))));

        Assert.Empty(PackageScanner.FindStalePivotCaches(path, Refresh().IsStale));
    }

    [Fact]
    public void Kazdy_stary_cache_jest_wymieniony_osobno()
    {
        string path = Package(
            (Part, Definition(Start.AddDays(-1))),
            ("xl/pivotCache/pivotCacheDefinition2.xml", Definition(Start.AddSeconds(1))),
            ("xl/pivotCache/pivotCacheDefinition3.xml", Definition(Start.AddDays(-2))));

        List<string> stale = PackageScanner.FindStalePivotCaches(path, Refresh().IsStale);

        Assert.Equal(2, stale.Count);
        Assert.Contains(stale, entry => entry.Contains("pivotCacheDefinition1.xml"));
        Assert.Contains(stale, entry => entry.Contains("pivotCacheDefinition3.xml"));
    }

    public void Dispose()
    {
        foreach (string file in _createdFiles)
        {
            try { File.Delete(file); } catch { }
        }
    }
}
