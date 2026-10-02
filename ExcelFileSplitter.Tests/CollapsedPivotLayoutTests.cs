using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class CollapsedPivotLayoutTests : IDisposable
{
    private readonly List<string> _createdFiles = [];

    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelations = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    private static readonly HashSet<string> Narrowed = new(StringComparer.OrdinalIgnoreCase) { "dim_Klienci", "fact_Sprzedaz" };

    private static readonly PivotTableRef Lista = new("PT_Lista", "Klienci");

    private static string Field(string name, string? caption, params string[] members) =>
        $"<cacheField name=\"{name}\"{(caption is null ? "" : $" caption=\"{caption}\"")} numFmtId=\"0\">" +
        $"<sharedItems>{string.Concat(members.Select(member => $"<s v=\"{member}\"/>"))}</sharedItems>" +
        $"<extLst><ext uri=\"{{4F2E5C28-24EA-4eb8-9CBF-B6C8F9C3D259}}\" xmlns:x15=\"{X15}\"><x15:cachedUniqueNames>" +
        string.Concat(members.Select((member, index) =>
            $"<x15:cachedUniqueName index=\"{index}\" name=\"{UniqueName(name, member)}\"/>")) +
        "</x15:cachedUniqueNames></ext></extLst></cacheField>";

    private static string UniqueName(string fieldName, string member) =>
        fieldName[..fieldName.LastIndexOf(".[", StringComparison.Ordinal)] + $".&amp;[{member}]";

    private const string Measure = "<cacheField name=\"[Measures].[Obrót]\" numFmtId=\"0\" hierarchy=\"80\" level=\"32767\"/>";

    private static string Cache(string connectionId, params string[] fields) =>
        $"<pivotCacheDefinition xmlns=\"{Main}\"><cacheSource type=\"external\" connectionId=\"{connectionId}\"/>" +
        $"<cacheFields count=\"{fields.Length}\">{string.Concat(fields)}</cacheFields></pivotCacheDefinition>";

    private static string RowField(params string[] items) =>
        $"<pivotField axis=\"axisRow\" compact=\"0\" outline=\"0\" showAll=\"0\" defaultAttributeDrillState=\"1\">" +
        $"<items count=\"{items.Length}\">{string.Concat(items)}</items></pivotField>";

    private const string DataField = "<pivotField dataField=\"1\" compact=\"0\" outline=\"0\" showAll=\"0\"/>";

    private static string Pivot(int[] rows, params string[] pivotFields) =>
        $"<pivotTableDefinition xmlns=\"{Main}\" name=\"PT_Lista\" cacheId=\"1\">" +
        "<location ref=\"I24:R30\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"2\"/>" +
        $"<pivotFields count=\"{pivotFields.Length}\">{string.Concat(pivotFields)}</pivotFields>" +
        $"<rowFields count=\"{rows.Length}\">{string.Concat(rows.Select(row => $"<field x=\"{row}\"/>"))}</rowFields>" +
        "</pivotTableDefinition>";

    private static readonly string CustomersCache = Cache("9",
        Field("[dim_Klienci].[Status].[Status]", "Status roczny", "Nowy", "Utracony"),
        Measure,
        Field("[dim_Klienci].[Grupa].[Grupa]", null, "BEL-POL", "MRÓWKA", "OBI"),
        Field("[dim_Klienci].[Klient].[Klient]", null, "ATREUS", "FENTIS", "KLARIS", "SOLEI"));

    private static CollapsedPivot? Read(string pivot, string cache) =>
        CollapsedPivotLayout.Read(Lista, XElement.Parse(pivot), XElement.Parse(cache), Narrowed);

    [Fact]
    public void Zwinieta_tabela_nad_polami_z_tabeli_zawezanej_jest_znaleziona()
    {
        string pivot = Pivot([0, 2, 3],
            RowField("<item x=\"0\" e=\"0\"/>", "<item x=\"1\" e=\"0\"/>"),
            DataField,
            RowField("<item x=\"0\"/>", "<item x=\"1\" e=\"0\"/>", "<item x=\"2\"/>"),
            RowField("<item x=\"0\"/>", "<item x=\"1\"/>", "<item x=\"2\"/>", "<item x=\"3\"/>"));

        CollapsedPivot found = Assert.IsType<CollapsedPivot>(Read(pivot, CustomersCache));

        Assert.Equal(Lista, found.Table);
        Assert.Equal(["dim_Klienci"], found.HiddenTables);
        Assert.Collection(found.Fields,
            status =>
            {
                Assert.Equal("[dim_Klienci].[Status].[Status]", status.Name);
                Assert.Equal("Status roczny", status.Caption);
                Assert.True(status.AllCollapsed);
                Assert.Equal(["[dim_Klienci].[Status].&[Nowy]", "[dim_Klienci].[Status].&[Utracony]"], status.CollapsedItems);
            },
            group =>
            {
                Assert.Equal("Grupa", group.Caption);
                Assert.False(group.AllCollapsed);
                Assert.Equal(["[dim_Klienci].[Grupa].&[MRÓWKA]"], group.CollapsedItems);
            });
        Assert.Equal(
            "tabela przestawna 'PT_Lista' (arkusz 'Klienci') ma zwinięte pozycje w polach: Status roczny (2 z 2), " +
            "Grupa (1 z 3) - pod nimi pamięta dane z 'dim_Klienci'",
            found.Describe());
    }

    [Fact]
    public void Tabela_rozwinieta_do_konca_nie_wymaga_rozwijania()
    {
        string pivot = Pivot([0, 3],
            RowField("<item x=\"0\"/>", "<item x=\"1\"/>"),
            DataField,
            DataField,
            RowField("<item x=\"0\"/>", "<item x=\"1\"/>", "<item x=\"2\"/>", "<item x=\"3\"/>"));

        Assert.Null(Read(pivot, CustomersCache));
    }

    [Fact]
    public void Zwiniete_pozycje_nad_tabela_niezawezana_nie_sa_ruszane()
    {
        string cache = Cache("9",
            Field("[dim_Klienci].[Klient].[Klient]", null, "ATREUS", "SOLEI"),
            Field("[dim_Produkty].[Produkt].[Produkt]", null, "Klej", "Fuga"));
        string pivot = Pivot([0, 1],
            RowField("<item x=\"0\" e=\"0\"/>", "<item x=\"1\" e=\"0\"/>"),
            RowField("<item x=\"0\"/>", "<item x=\"1\"/>"));

        Assert.Null(Read(pivot, cache));
    }

    [Fact]
    public void Pozycje_sumy_nie_sa_liczone_a_nazwa_z_wartosci_wystarcza()
    {
        string cache = Cache("9",
            "<cacheField name=\"[dim_Klienci].[Grupa].[Grupa]\" numFmtId=\"0\"><sharedItems>" +
            "<s v=\"[dim_Klienci].[Grupa].&amp;[OBI]\" c=\"OBI\"/><s v=\"[dim_Klienci].[Grupa].&amp;[BRICO]\" c=\"BRICO\"/>" +
            "</sharedItems></cacheField>",
            Field("[dim_Klienci].[Klient].[Klient]", null, "ATREUS"));
        string pivot = Pivot([0, 1],
            RowField("<item x=\"0\" e=\"0\"/>", "<item x=\"1\" e=\"0\"/>", "<item t=\"default\"/>"),
            RowField("<item x=\"0\"/>"));

        CollapsedField group = Assert.Single(Assert.IsType<CollapsedPivot>(Read(pivot, cache)).Fields);

        Assert.True(group.AllCollapsed);
        Assert.Equal(["[dim_Klienci].[Grupa].&[OBI]", "[dim_Klienci].[Grupa].&[BRICO]"], group.CollapsedItems);
    }

    [Theory]
    [InlineData("[dim_Klienci].[Nazwa Klienta]", "dim_Klienci")]
    [InlineData("[dim_Klienci].[Nazwa Klienta].[Nazwa Klienta]", "dim_Klienci")]
    [InlineData("[fakt]]y].[Klient]", "fakt]y")]
    [InlineData("[Measures].[Obrót]", null)]
    [InlineData("[Zestaw klientów]", null)]
    [InlineData("Values", null)]
    [InlineData("", null)]
    public void Tabela_pola_z_jego_nazwy_w_modelu(string name, string? table)
    {
        Assert.Equal(table, CollapsedPivotLayout.TableOf(name));
    }

    [Theory]
    [InlineData(new[] { "dim_Klienci", "dim_Klienci", "dim_Klienci" }, 2)]
    [InlineData(new[] { "dim_Produkty", "DIM_KLIENCI", "dim_Produkty" }, 1)]
    [InlineData(new[] { "dim_Klienci", "dim_Produkty" }, 0)]
    [InlineData(new[] { "dim_Klienci" }, 0)]
    [InlineData(new string[0], 0)]
    public void Liczba_poziomow_ktore_moga_ukryc_dane_odbiorcow(string[] tables, int levels)
    {
        Assert.Equal(levels, CollapsedPivotLayout.LevelsHidingNarrowed(tables, Narrowed));
    }

    [Fact]
    public void Zawezane_sa_wymiar_podzialu_i_zapytania_z_filtrem()
    {
        HashSet<string> narrowed = SplitConfig.CreateSample().PowerPivot.NarrowedTables();

        Assert.Equal(["dim_KlienciV2", "fact_Zamówienia"], narrowed.Order());
        Assert.Contains("DIM_KLIENCIV2", narrowed);
    }

    [Fact]
    public void Plik_czytany_jest_przez_arkusz_i_tylko_dla_cache_modelu()
    {
        string pivot = Pivot([0, 3],
            RowField("<item x=\"0\" e=\"0\"/>", "<item x=\"1\"/>"),
            DataField,
            DataField,
            RowField("<item x=\"0\"/>", "<item x=\"1\"/>", "<item x=\"2\"/>", "<item x=\"3\"/>"));

        CollapsedPivot found = Assert.Single(CollapsedPivotLayout.Find(Package(pivot, CustomersCache), Narrowed));
        Assert.Equal(Lista, found.Table);
        Assert.Equal("Status roczny (1 z 2)", Assert.Single(found.Fields).ToString());

        Assert.Empty(CollapsedPivotLayout.Find(Package(pivot, CustomersCache.Replace("connectionId=\"9\"", "connectionId=\"3\"")), Narrowed));
    }

    private string Package(string pivot, string cache)
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.zip");
        _createdFiles.Add(path);

        (string Part, string Xml)[] parts =
        [
            ("xl/workbook.xml",
                $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Relations}\"><sheets>" +
                "<sheet name=\"Klienci\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"),
            ("xl/_rels/workbook.xml.rels", Rels(Relationship("rId1", "worksheet", "worksheets/sheet1.xml"))),
            ("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{Main}\"/>"),
            ("xl/worksheets/_rels/sheet1.xml.rels", Rels(Relationship("rId1", "pivotTable", "../pivotTables/pivotTable1.xml"))),
            ("xl/pivotTables/pivotTable1.xml", pivot),
            ("xl/pivotTables/_rels/pivotTable1.xml.rels",
                Rels(Relationship("rId1", "pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml"))),
            ("xl/pivotCache/pivotCacheDefinition1.xml", cache),
            ("xl/connections.xml",
                $"<connections xmlns=\"{Main}\">" +
                "<connection id=\"3\" name=\"Hurtownia\" type=\"5\"><dbPr connection=\"x\" command=\"y\"/></connection>" +
                "<connection id=\"9\" name=\"ThisWorkbookDataModel\" type=\"5\"><dbPr connection=\"x\" command=\"Model\"/>" +
                "</connection></connections>"),
        ];

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

    public void Dispose()
    {
        foreach (string path in _createdFiles)
        {
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
            }
        }
    }
}
