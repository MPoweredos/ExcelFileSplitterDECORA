using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class PivotCacheMembersTests : IDisposable
{
    private static readonly CultureInfo Polski = CultureInfo.GetCultureInfo("pl-PL");

    private readonly List<string> _createdFiles = [];

    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelations = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

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

    private static string Relationship(string id, string type, string target) =>
        $"<Relationship Id=\"{id}\" Type=\"{Relations}/{type}\" Target=\"{target}\"/>";

    private static string Rels(params string[] relationships) =>
        $"<Relationships xmlns=\"{PackageRelations}\">{string.Concat(relationships)}</Relationships>";

    private static (string, string)[] Workbook(params (string Part, string Xml)[] caches) =>
    [
        ("xl/workbook.xml",
            $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Relations}\"><sheets>" +
            "<sheet name=\"Klienci\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"),
        ("xl/_rels/workbook.xml.rels", Rels(Relationship("rId1", "worksheet", "worksheets/sheet1.xml"))),
        ("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{Main}\"/>"),
        ("xl/worksheets/_rels/sheet1.xml.rels", Rels(Relationship("rId1", "pivotTable", "../pivotTables/pivotTable1.xml"))),
        ("xl/pivotTables/pivotTable1.xml", $"<pivotTableDefinition xmlns=\"{Main}\" name=\"PT_Lista\" cacheId=\"1\"/>"),
        ("xl/pivotTables/_rels/pivotTable1.xml.rels",
            Rels(Relationship("rId1", "pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml"))),
        ("xl/connections.xml",
            $"<connections xmlns=\"{Main}\">" +
            "<connection id=\"3\" name=\"Hurtownia\" type=\"5\"><dbPr connection=\"x\" command=\"y\"/></connection>" +
            "<connection id=\"9\" name=\"ThisWorkbookDataModel\" type=\"5\"><dbPr connection=\"x\" command=\"Model\"/>" +
            $"<extLst><ext uri=\"{{DE250136-89BD-433C-8126-D09CA5730AF9}}\" xmlns:x15=\"{X15}\">" +
            "<x15:connection id=\"\" model=\"1\"/></ext></extLst></connection></connections>"),
        .. caches,
    ];

    private static string Cache(string source, params string[] fields) =>
        $"<pivotCacheDefinition xmlns=\"{Main}\" xmlns:r=\"{Relations}\" refreshedDate=\"46295.5\">{source}" +
        $"<cacheFields count=\"{fields.Length}\">{string.Concat(fields)}</cacheFields></pivotCacheDefinition>";

    private const string ModelSource = "<cacheSource type=\"external\" connectionId=\"9\"/>";

    private static string Field(string name, string items, params string[] uniqueNames) =>
        $"<cacheField name=\"{name}\" numFmtId=\"0\"><sharedItems>{items}</sharedItems>" +
        (uniqueNames.Length == 0
            ? ""
            : $"<extLst><ext uri=\"{{4F2E5C28-24EA-4eb8-9CBF-B6C8F9C3D259}}\" xmlns:x15=\"{X15}\"><x15:cachedUniqueNames>" +
              string.Concat(uniqueNames.Select((uniqueName, index) =>
                  $"<x15:cachedUniqueName index=\"{index}\" name=\"{uniqueName}\"/>")) +
              "</x15:cachedUniqueNames></ext></extLst>") +
        "</cacheField>";

    private string ReportWithStaleGroups() => Package(Workbook(
        ("xl/pivotCache/pivotCacheDefinition1.xml", Cache(ModelSource,
            Field("[dim_Klienci].[Grupa].[Grupa]", "<s v=\"BEL-POL\"/><s v=\"MRÓWKA\"/>",
                "[dim_Klienci].[Grupa].&amp;[BEL-POL]", "[dim_Klienci].[Grupa].&amp;[MRÓWKA]"),
            "<cacheField name=\"[Measures].[Obrót]\" numFmtId=\"0\" hierarchy=\"80\" level=\"32767\"/>",
            Field("[dim_Kalendarz].[RokOffset].[RokOffset]", "<s v=\"[dim_Kalendarz].[RokOffset].&amp;[0]\" c=\"0\"/>"),
            Field("[dim_Kalendarz].[Rok].[Rok]", "<n v=\"2026\"/>", "[dim_Kalendarz].[Rok].&amp;[2026]"))),
        ("xl/pivotCache/pivotCacheDefinition2.xml", Cache("<cacheSource type=\"external\" connectionId=\"3\"/>",
            Field("[Klienci].[Grupa].[Grupa]", "<s v=\"CASTORAMA\"/>"))),
        ("xl/pivotCache/pivotCacheDefinition3.xml", Cache(
            "<cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:B9\" sheet=\"Dane\"/></cacheSource>",
            Field("Grupa", "<s v=\"CASTORAMA\"/>")))));

    [Fact]
    public void Czyta_pozycje_tabel_przestawnych_modelu_razem_z_wlascicielem()
    {
        List<CachedModelField> fields = PivotCacheMembers.Read(ReportWithStaleGroups());

        Assert.Equal(["Grupa", "RokOffset", "Rok"], fields.Select(field => field.Column));
        Assert.All(fields, field => Assert.Equal("'PT_Lista' (arkusz 'Klienci')", field.Owner));

        CachedModelField groups = fields[0];
        Assert.Equal("dim_Klienci", groups.Table);
        Assert.Equal([("BEL-POL", "BEL-POL"), ("MRÓWKA", "MRÓWKA")], groups.Members.Select(member => (member.Caption, member.Key)));
        Assert.Equal(new CachedMember("0", "0"), fields[1].Members.Single());
        Assert.Equal(new CachedMember("2026", "2026"), fields[2].Members.Single());
    }

    [Fact]
    public void Pozycja_spoza_modelu_odbiorcy_jest_zglaszana()
    {
        CachedModelField groups = PivotCacheMembers.Read(ReportWithStaleGroups())[0];

        List<string> outside = PivotCacheMembers.OutsideModel(groups, new ModelValues(["MRÓWKA", "BRICOMARCHE"], Polski));

        Assert.Equal(["BEL-POL"], outside);
    }

    [Fact]
    public void Cache_bez_polaczenia_z_modelem_nie_jest_czytany()
    {
        string path = Package(Workbook(
            ("xl/pivotCache/pivotCacheDefinition1.xml", Cache("<cacheSource type=\"external\" connectionId=\"3\"/>",
                Field("[Klienci].[Grupa].[Grupa]", "<s v=\"CASTORAMA\"/>")))));

        Assert.Empty(PivotCacheMembers.Read(path));
    }

    [Fact]
    public void Klucz_liczby_i_daty_pasuje_mimo_innego_zapisu_w_modelu()
    {
        var values = new ModelValues(["12,5", "30.09.2026 00:00:00"], Polski);

        Assert.True(values.Contains(new CachedMember("12,50 zł", "12.5")));
        Assert.True(values.Contains(new CachedMember("wrzesień 2026", "2026-09-30T00:00:00")));
        Assert.False(values.Contains(new CachedMember("CASTORAMA", "CASTORAMA")));
    }

    [Fact]
    public void Numer_z_klucza_spoza_modelu_jest_zglaszany()
    {
        var values = new ModelValues(["108535", "109975"], Polski);

        Assert.True(values.Contains(new CachedMember("108535", "108535")));
        Assert.False(values.Contains(new CachedMember("999999", "999999")));
    }

    [Fact]
    public void Podpis_liczby_bez_klucza_nie_blokuje_bo_nie_da_sie_go_porownac()
    {
        var values = new ModelValues(["MRÓWKA"], Polski);

        Assert.True(values.Contains(new CachedMember("1234,5", null)));
        Assert.True(values.Contains(new CachedMember("2026-09-30", null)));
        Assert.False(values.Contains(new CachedMember("CASTORAMA", null)));
    }

    [Fact]
    public void Puste_i_syntetyczne_pozycje_nie_blokuja()
    {
        var values = new ModelValues(["MRÓWKA"], Polski);

        Assert.True(values.Contains(new CachedMember("", "")));
        Assert.True(values.Contains(new CachedMember("(blank)", null)));
        Assert.True(values.Contains(new CachedMember("(puste)", null)));
    }

    [Fact]
    public void Wielkosc_liter_i_spacje_na_brzegach_nie_maja_znaczenia()
    {
        var values = new ModelValues([" Bochnia"], Polski);

        Assert.True(values.Contains(new CachedMember("BOCHNIA ", "BOCHNIA ")));
    }

    [Theory]
    [InlineData("[t].[h].&[BEL-POL]", "BEL-POL")]
    [InlineData("[t].[h].&[A]]B]", "A]B")]
    [InlineData("[T].[H].&[x]", "x")]
    [InlineData("[t].[h].&[2026]&[9]", null)]
    [InlineData("[t].[inna].&[x]", null)]
    [InlineData("BEL-POL", null)]
    public void Klucz_jest_czytany_tylko_z_prostej_nazwy_pozycji(string uniqueName, string? expected)
    {
        Assert.Equal(expected, PivotCacheMembers.KeyOf(uniqueName, "[t].[h]"));
    }

    [Fact]
    public void Nazwa_pola_jest_dzielona_z_uwzglednieniem_nawiasow()
    {
        Assert.Equal(["dim_K", "Nazwa]x", "Nazwa]x"], PivotCacheMembers.BracketParts("[dim_K].[Nazwa]]x].[Nazwa]]x]"));
        Assert.Empty(PivotCacheMembers.BracketParts("Grupa"));
        Assert.Empty(PivotCacheMembers.BracketParts("[a].[b"));
    }

    [Fact]
    public void Pamietane_pozycje_spoza_modelu_blokuja_plik()
    {
        var result = new VerificationResult(
            @"C:\out\Raport.xlsm", ["PH MINI DIY GDAŃSK"], ["PH MINI DIY GDAŃSK"], [], [], [], [], [], 361, null)
        {
            StalePivotMembers = ["'PT_02_Lista' (arkusz '02_Customers1') [Nazwa Grupy Sprzedawców] -> 5: BEL-POL, CASTORAMA"],
        };

        Assert.False(result.Ok);
        Assert.Contains("pamiętają pozycje spoza modelu", result.Summary());
        Assert.Contains("PT_02_Lista", result.Summary());
        Assert.True((result with { StalePivotMembers = [] }).Ok);
    }

    public void Dispose()
    {
        foreach (string file in _createdFiles)
        {
            try { File.Delete(file); } catch { }
        }
    }
}
