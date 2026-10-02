using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelFileSplitter.Core;

public sealed partial class SplitConfig
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,

        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    public static SplitConfig Load(string path)
    {
        string json = File.ReadAllText(path);
        SplitConfig config = JsonSerializer.Deserialize<SplitConfig>(json, JsonOptions)
                             ?? throw new InvalidOperationException($"Nie udało się wczytać konfiguracji z {path}");

        config.ReplaceNullSections();
        return config;
    }

    private void ReplaceNullSections()
    {
        Worksheet ??= new WorksheetOptions();
        PowerPivot ??= new PowerPivotOptions();
        Hardening ??= new HardeningOptions();
    }

    public void Save(string path) => File.WriteAllText(path, JsonSerializer.Serialize(this, JsonOptions));

    public void SaveTemplate(string path) =>
        File.WriteAllText(path, TemplateHelp + JsonSerializer.Serialize(this, JsonOptions));

    private const string TemplateHelp = """
        // =========================================================================
        // Ustawienia podziału. Wypełnij i wczytaj przyciskiem "Wczytaj ustawienia".
        // Komentarze (linie z //) wolno zostawić - przy wczytywaniu są pomijane.
        // =========================================================================
        //
        // KTO DOSTAJE KTÓRE DANE - wypełnij JEDNO z dwojga, nigdy oba naraz:
        //
        //   "Values": ["Anna Kowalska", "Piotr Nowak"]
        //       Osobny plik dla każdej wymienionej wartości kolumny podziału.
        //       Puste Values = plik dla KAŻDEJ wartości występującej w danych.
        //
        //   "Groups": [
        //     {
        //       "Name": "Zespół Północ",
        //       "Values": ["Anna Kowalska", "Piotr Nowak"],
        //       "SheetsToKeep": ["Raport Północ"]
        //     },
        //     {
        //       "Name": "Anna Zielińska",
        //       "Values": ["Anna Zielińska"],
        //       "SheetsToKeep": []
        //     }
        //   ]
        //       Pliki zbiorcze: jedna grupa = jeden plik z danymi wszystkich swoich wartości.
        //       Name trafia do nazwy pliku w miejsce {name} / {value} z FileNameTemplate.
        //       SheetsToKeep w grupie to arkusze PONAD wspólną listę SheetsToKeep poniżej,
        //       a nie zamiast niej.
        //
        // Grupy da się też poklikać w oknie (krok "Grupy i arkusze") - wtedy tej sekcji
        // nie trzeba wypełniać ręcznie.
        //
        // SEKCJE TRYBÓW: "Worksheet" i "PowerPivot". Nazywają się tak samo jak wartości
        // "Mode", więc obowiązuje Cię ta, której nazwa zgadza się z Twoim Mode - drugą
        // możesz pominąć w całości, nic z niej nie jest czytane.
        //
        // Poniższe ustawienia siedzą wewnątrz "PowerPivot": { ... }.
        //
        // CACHE - cztery różne ustawienia o podobnej nazwie:
        //
        //   "UseCache"        czy PODZIAŁ ma czytać dane z plików CSV zamiast ze źródeł.
        //                     To jedyny przełącznik, którym sterujesz na co dzień. W oknie
        //                     ustawia go krok "Cache (etap 1)" i to jego stan jest brany
        //                     pod uwagę - przy przejściu dalej nadpisuje wartość stąd.
        //   "CacheFolder"     gdzie leżą te pliki CSV.
        //   "CacheTables"     które tabele modelu zrzucać do CSV. Puste = wszystkie.
        //   "Queries[].Cache" czy do tego zapytania podpiąć przełącznik czytania z CSV.
        //                     Robi to jednorazowo etap "przygotowanie"; zostaw true dla
        //                     zapytań sięgających do plików źródłowych.
        //
        //   "CacheSelfCheck"     po zrzucie zbuduj model z CSV i porównaj z oryginałem.
        //                        Kosztuje jedno odświeżenie, a łapie rozjazd przed podziałem.
        //   "RefreshPivotCaches" dodatkowy przebieg po cache'ach tabel przestawnych.
        //                        Bardzo wolny, RefreshAll odświeża je samo - włącz tylko,
        //                        gdy w gotowych plikach zobaczysz nieaktualne przestawne.
        // =========================================================================

        """;

    public static SplitConfig CreateWorksheetSample() => new()
    {
        Mode = SplitMode.Worksheet,
        SourceWorkbook = @"C:\Raporty\Sprzedaż.xlsx",
        OutputFolder = @"C:\Raporty\out",
        Worksheet = new WorksheetOptions { Table = "tblSprzedaz", Sheet = "", HeaderRow = 1 },
        SplitColumn = "Nazwa Opiekuna Klienta",
        Values = [],
        SheetsToKeep = ["Dane", "Dashboard"],
        FileNameTemplate = "Sprzedaż_{name}.xlsx",
    };

    public static SplitConfig CreateSample() => new()
    {
        SourceWorkbook = @"C:\Raporty\Raport sprzedaży.xlsx",
        OutputFolder = @"C:\Raporty\out",
        SplitColumn = "Nazwa Opiekuna Klienta",
        Values = [],
        SheetsToKeep = ["Dashboard", "Analiza klientów"],
        FileNameTemplate = "Raport_{name}.xlsx",
        PowerPivot = new PowerPivotOptions
        {
            CacheFolder = @"C:\Raporty\cache",
            UseCache = false,
            SplitTable = "dim_KlienciV2",
            Queries =
            [
                new QuerySpec { Name = "dim_KlienciV2", FilterColumn = "Nazwa Opiekuna Klienta", Cache = true },
                new QuerySpec
                {
                    Name = "fact_Zamówienia",
                    FilterColumn = "Odbiorca materiałów",
                    KeyFrom = new KeyFrom { Query = "dim_KlienciV2", Column = "Identyfikator" },
                    Cache = true,
                },
                new QuerySpec { Name = "dim_Produkty", Cache = true },
            ],
        },
    };
}
