using System.IO.Compression;
using System.Text;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class QueryRefreshLogTests : IDisposable
{
    private readonly List<string> _createdFiles = [];

    private static readonly DateTime SplitStartedUtc = new(2026, 9, 30, 17, 51, 0, DateTimeKind.Utc);

    private string Package(params (string Part, byte[] Content)[] parts)
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.zip");
        _createdFiles.Add(path);

        using var package = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach ((string part, byte[] content) in parts)
        {
            using Stream stream = package.CreateEntry(part).Open();
            stream.Write(content, 0, content.Length);
        }
        return path;
    }

    private static byte[] WithLength(byte[] content) => [.. BitConverter.GetBytes(content.Length), .. content];

    private static string Item(string path, params (string Type, string Value)[] entries) =>
        $"<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>{path}</ItemPath></ItemLocation><StableEntries>" +
        string.Concat(entries.Select(entry => $"<Entry Type=\"{entry.Type}\" Value=\"{entry.Value}\" />")) +
        "</StableEntries></Item>";

    private static byte[] Mashup(params string[] items)
    {
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LocalPackageMetadataFile " +
                     "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                     "<Items><Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries /></Item>" +
                     string.Concat(items) + "</Items></LocalPackageMetadataFile>";
        byte[] metadata = [.. BitConverter.GetBytes(0), .. WithLength([.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes(xml)]), .. WithLength([])];

        return
        [
            .. BitConverter.GetBytes(0),
            .. WithLength(Encoding.ASCII.GetBytes("PK-udawana-paczka")),
            .. WithLength(Encoding.UTF8.GetBytes("<PermissionList><FirewallEnabled>true</FirewallEnabled></PermissionList>")),
            .. WithLength(metadata),
            .. WithLength([]),
        ];
    }

    private static byte[] CustomXml(byte[] mashup) =>
    [
        .. Encoding.Unicode.GetPreamble(),
        .. Encoding.Unicode.GetBytes("<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup " +
                                     "xmlns=\"http://schemas.microsoft.com/DataMashup\">" +
                                     Convert.ToBase64String(mashup) + "</DataMashup>"),
    ];

    private string Report() => Package(
        ("customXml/item1.xml", Encoding.UTF8.GetBytes("<Gemini xmlns=\"http://gemini/pivotcustomization/TableOrder\"/>")),
        ("customXml/item59.xml", CustomXml(Mashup(
            Item("Section1/fact_Zam%C3%B3wienia", ("FillLastUpdated", "d2026-08-31T10:36:09.4030441Z"), ("FillCount", "l991070")),
            Item("Section1/dim_KlienciV2", ("FillLastUpdated", "d2026-09-30T17:52:34.6620252Z"), ("FillCount", "l145")),
            Item("Section1/Transform%20File", ("FillLastUpdated", "d2026-07-29T08:17:55.6087489Z"))))));

    [Fact]
    public void Czyta_date_odswiezenia_i_liczbe_wierszy_zapytan()
    {
        Dictionary<string, QueryFill> fills = QueryRefreshLog.Read(Report())!;

        Assert.Equal(new DateTime(2026, 8, 31, 10, 36, 9, DateTimeKind.Utc), fills["fact_Zamówienia"].LastFilledUtc!.Value.AddTicks(-4030441));
        Assert.Equal(991_070, fills["fact_Zamówienia"].Rows);
        Assert.Equal(145, fills["dim_KlienciV2"].Rows);
        Assert.Null(fills["Transform File"].Rows);
    }

    [Fact]
    public void Zapytanie_nieodswiezone_przy_podziale_jest_zglaszane()
    {
        Dictionary<string, QueryFill> fills = QueryRefreshLog.Read(Report())!;

        List<QueryFill> stale = QueryRefreshLog.NotRefreshed(fills, ["dim_KlienciV2", "fact_Zamówienia"], SplitStartedUtc);

        Assert.Equal(["fact_Zamówienia"], stale.Select(fill => fill.Query));
    }

    [Fact]
    public void Zapytanie_bez_zapisu_odswiezenia_tez_jest_zglaszane()
    {
        Dictionary<string, QueryFill> fills = QueryRefreshLog.Read(Report())!;

        QueryFill stale = Assert.Single(QueryRefreshLog.NotRefreshed(fills, ["dim_Produkty"], SplitStartedUtc));

        Assert.Null(stale.LastFilledUtc);
    }

    [Fact]
    public void Odswiezenie_w_tej_samej_sekundzie_co_start_nie_jest_alarmem()
    {
        Dictionary<string, QueryFill> fills = QueryRefreshLog.Read(Report())!;

        Assert.Empty(QueryRefreshLog.NotRefreshed(fills, ["dim_KlienciV2"], new DateTime(2026, 9, 30, 17, 52, 34, 900, DateTimeKind.Utc)));
    }

    [Fact]
    public void Plik_bez_pakietu_zapytan_nie_ma_zapisu_odswiezenia()
    {
        string path = Package(("customXml/item1.xml", Encoding.UTF8.GetBytes("<Gemini xmlns=\"http://gemini/pivotcustomization/TableOrder\"/>")));

        Assert.Null(QueryRefreshLog.Read(path));
    }

    [Fact]
    public void Uszkodzony_pakiet_zapytan_to_blad_a_nie_zgoda()
    {
        Assert.Throws<InvalidDataException>(() => QueryRefreshLog.ReadMetadata([0, 0, 0, 0, 200, 0, 0, 0, 1, 2]));
    }

    public void Dispose()
    {
        foreach (string file in _createdFiles)
        {
            try { File.Delete(file); } catch { }
        }
    }
}
