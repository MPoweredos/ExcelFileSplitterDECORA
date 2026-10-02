using System.IO.Compression;
using System.Text;
using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class PackageScannerTests : IDisposable
{
    private const string Tabela = "dim_Klienci";
    private const string Kolumna = "Opiekun";

    private readonly List<string> _createdFiles = [];

    private string Package(string part, byte[] content)
    {
        string path = Path.Combine(Path.GetTempPath(), $"efs-test-{Guid.NewGuid():N}.zip");
        _createdFiles.Add(path);

        using (var package = ZipFile.Open(path, ZipArchiveMode.Create))
        using (var stream = package.CreateEntry(part).Open())
        {
            stream.Write(content, 0, content.Length);
        }
        return path;
    }

    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
    private static byte[] Utf16(string text) => Encoding.Unicode.GetBytes(text);

    [Fact]
    public void Obrazy_sa_pomijane()
    {
        string path = Package("xl/media/image1.png", Utf8("ANNA NOWAK"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Empty(hits);
    }

    [Fact]
    public void Kwalifikowane_odwolanie_OLAP_to_jednoznaczny_wyciek()
    {
        string path = Package("xl/pivotCache/pivotCacheDefinition1.xml",
            Utf8($"<x><item c=\"[{Tabela}].[{Kolumna}].&[ANNA NOWAK]\"/></x>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Single(hits);
        Assert.True(hits[0].IsQualifiedReference);
        Assert.Equal("ANNA NOWAK", hits[0].Value);
        Assert.Contains("pivotCacheDefinition1.xml", hits[0].PartName);
    }

    [Fact]
    public void Ampersand_zapisany_jako_encja_XML_tez_jest_rozpoznawany()
    {
        string path = Package("xl/charts/chart1.xml",
            Utf8($"<c:pt><c:v>[{Tabela}].[{Kolumna}].&amp;[ANNA NOWAK]</c:v></c:pt>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Single(hits);
        Assert.True(hits[0].IsQualifiedReference);
    }

    [Fact]
    public void Sama_nazwa_bez_kontekstu_to_wzmianka_a_nie_wyciek()
    {
        string path = Package("xl/worksheets/sheet1.xml", Utf8("<x><c t=\"s\"><v>ANNA NOWAK</v></c></x>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Single(hits);
        Assert.False(hits[0].IsQualifiedReference);
    }

    [Fact]
    public void Tresc_w_UTF16_jest_przeszukiwana_tak_samo()
    {
        string path = Package("customXml/item1.xml", Utf16("<DataModel><Opiekun>ANNA NOWAK</Opiekun></DataModel>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Single(hits);
        Assert.Equal("ANNA NOWAK", hits[0].Value);
    }

    [Fact]
    public void Wielkosc_liter_nie_ukrywa_wycieku()
    {
        string path = Package("xl/worksheets/sheet1.xml", Utf8("<x><v>anna nowak</v></x>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Single(hits);
    }

    [Fact]
    public void Czysta_paczka_nie_daje_trafien()
    {
        string path = Package("xl/worksheets/sheet1.xml", Utf8("<x><v>JAN KOWALSKI</v></x>"));

        var hits = PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]);

        Assert.Empty(hits);
    }

    [Fact]
    public void Fragment_innego_slowa_to_nie_wyciek()
    {
        string path = Package("xl/pivotCache/pivotCacheDefinition1.xml",
            Utf8("<sharedItems><s v=\"MRÓWKA KOBIERZYCE\"/><s v=\"LISTWA MOBILNA\"/></sharedItems>"));

        Assert.Empty(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Wartosc_jako_osobne_slowo_w_dluzszym_tekscie_to_wyciek()
    {
        string path = Package("xl/charts/chart1.xml", Utf8("<c:tx><a:t>Sprzedaż OBI 2024</a:t></c:tx>"));

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Trafienie_pokazuje_otoczenie_w_pliku()
    {
        string path = Package("xl/pivotCache/pivotCacheDefinition1.xml",
            Utf8("<sharedItems><s v=\"LEROY MERLIN\"/><s v=\"OBI\"/></sharedItems>"));

        ScanHit hit = Assert.Single(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));

        Assert.Contains("<s v=\"OBI\"/>", hit.Context);
        Assert.Contains(hit.Context!, hit.Describe());
    }

    [Fact]
    public void Wartosc_ze_znakiem_specjalnym_zapisanym_encja_XML_jest_znajdowana()
    {
        string path = Package("xl/sharedStrings.xml", Utf8("<sst><si><t>B&amp;Q</t></si></sst>"));

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["B&Q"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Wartosc_konczaca_sie_kropka_nie_wymaga_granicy_slowa_za_kropka()
    {
        string path = Package("xl/sharedStrings.xml", Utf8("<sst><si><t>PH SPEC.KRAKÓW</t></si></sst>"));

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["PH SPEC."], Tabela, [Kolumna]));
    }

    [Fact]
    public void Czesc_binarna_spoza_xlsb_nadal_jest_przeszukiwana_po_fragmencie()
    {
        string path = Package("xl/printerSettings/printerSettings1.bin", Utf16("MRÓWKA KOBIERZYCE"));

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    private static byte[] Record(int type, params byte[][] payload)
    {
        byte[] body = payload.SelectMany(part => part).ToArray();
        var bytes = new List<byte>();
        if (type < 0x80) bytes.Add((byte)type);
        else bytes.AddRange([(byte)((type & 0x7F) | 0x80), (byte)(type >> 7)]);

        int size = body.Length;
        do
        {
            byte current = (byte)(size & 0x7F);
            size >>= 7;
            bytes.Add(size > 0 ? (byte)(current | 0x80) : current);
        } while (size > 0);

        bytes.AddRange(body);
        return [.. bytes];
    }

    private static byte[] WideString(string text) => [.. BitConverter.GetBytes(text.Length), .. Utf16(text)];

    private static byte[] FormulaString(string text) => [0x17, .. BitConverter.GetBytes((ushort)text.Length), .. Utf16(text)];

    private static byte[] SharedStrings(params string[] texts) =>
    [
        .. Record(0x9F, BitConverter.GetBytes(texts.Length), BitConverter.GetBytes(texts.Length)),
        .. texts.SelectMany(text => Record(0x13, [0x00], WideString(text))),
        .. Record(0xA0),
    ];

    [Fact]
    public void Fragment_innego_slowa_w_xlsb_to_nie_wyciek()
    {
        string path = Package("xl/sharedStrings.bin",
            SharedStrings("POMPA OBIEGOWA", "UL. SOBIESKIEGO 12", "KOBIERZYCE", "LISTWA MOBILNA"));

        Assert.Empty(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Caly_tekst_w_xlsb_to_wyciek_niezaleznie_od_bajtu_parzystego_czy_nie()
    {
        string first = Package("xl/sharedStrings.bin", SharedStrings("OBI", "LEROY MERLIN"));
        string second = Package("xl/sharedStrings.bin", SharedStrings("LEROY MERLIN", "OBI"));

        Assert.Single(PackageScanner.FindForbiddenValues(first, ["OBI"], Tabela, [Kolumna]));
        Assert.Single(PackageScanner.FindForbiddenValues(second, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Bajty_za_tekstem_w_xlsb_czytane_jako_litera_nie_ukrywaja_wycieku()
    {
        string whole = Package("xl/worksheets/sheet1.bin", Record(0x2A, WideString("OBI"), BitConverter.GetBytes(0x41)));
        string ending = Package("xl/worksheets/sheet1.bin", Record(0x2A, WideString("SKLEP OBI"), BitConverter.GetBytes(0x113)));
        string multiline = Package("xl/worksheets/sheet1.bin", Record(0x2A, WideString("SKLEP\nOBI"), BitConverter.GetBytes(0x41)));

        Assert.Single(PackageScanner.FindForbiddenValues(whole, ["OBI"], Tabela, [Kolumna]));
        Assert.Single(PackageScanner.FindForbiddenValues(ending, ["OBI"], Tabela, [Kolumna]));
        Assert.Single(PackageScanner.FindForbiddenValues(multiline, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Szum_odczytany_jako_litera_innego_pisma_nie_ukrywa_wycieku_w_xlsb()
    {
        string path = Package("xl/worksheets/sheet1.bin", Record(0x2A, Utf16("X OBI"), [0x07, 0x0C]));

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Tekst_w_formule_xlsb_jest_znajdowany_takze_przed_litera_z_nastepnego_tokenu()
    {
        byte[] areaReference = [0x45, 0x00, 0x00, 0x00, 0x00];
        string shortText = Package("xl/worksheets/sheet1.bin", Record(0x08, FormulaString("OBI"), areaReference));
        string longText = Package("xl/worksheets/sheet1.bin",
            Record(0x08, FormulaString("SPRZEDAŻ W SIECIACH: BRICOMAN I OBI"), areaReference));
        string startsWith = Package("xl/worksheets/sheet1.bin",
            Record(0x08, FormulaString("OBI " + new string('X', 61)), areaReference));

        Assert.Single(PackageScanner.FindForbiddenValues(shortText, ["OBI"], Tabela, [Kolumna]));
        Assert.Single(PackageScanner.FindForbiddenValues(longText, ["OBI"], Tabela, [Kolumna]));
        Assert.Single(PackageScanner.FindForbiddenValues(startsWith, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Przypadkowe_bajty_liczb_w_xlsb_nie_sa_wyciekiem()
    {
        byte[] number = [0x00, 0x00, 0x6F, 0x62, 0x69, 0x21, 0x40, 0x40];
        string path = Package("xl/worksheets/sheet1.bin",
            Record(0x05, BitConverter.GetBytes(3), BitConverter.GetBytes(0), number));

        Assert.Empty(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Trafienie_w_xlsb_pokazuje_czytelne_otoczenie()
    {
        string path = Package("xl/sharedStrings.bin", SharedStrings("LEROY MERLIN", "SKLEP OBI 24H", "CASTORAMA"));

        ScanHit hit = Assert.Single(PackageScanner.FindForbiddenValues(path, ["OBI"], Tabela, [Kolumna]));

        Assert.Contains("SKLEP OBI 24H", hit.Context);
        Assert.DoesNotContain(hit.Context!, char.IsControl);
    }

    [Fact]
    public void Tekst_UTF16_od_nieparzystego_bajtu_w_czesci_binarnej_jest_znajdowany()
    {
        byte[] content = [0x07, .. Utf16("ANNA NOWAK")];
        string path = Package("xl/worksheets/sheet1.bin", content);

        Assert.Single(PackageScanner.FindForbiddenValues(path, ["ANNA NOWAK"], Tabela, [Kolumna]));
    }

    [Fact]
    public void Pusta_lista_wartosci_zakazanych_konczy_skan_od_razu()
    {
        string path = Package("xl/worksheets/sheet1.xml", Utf8("<x><v>ANNA NOWAK</v></x>"));

        Assert.Empty(PackageScanner.FindForbiddenValues(path, [], Tabela, [Kolumna]));
        Assert.Empty(PackageScanner.FindForbiddenValues(path, ["   "], Tabela, [Kolumna]));
    }

    public void Dispose()
    {
        foreach (string file in _createdFiles)
        {
            try { File.Delete(file); } catch { }
        }
    }
}
