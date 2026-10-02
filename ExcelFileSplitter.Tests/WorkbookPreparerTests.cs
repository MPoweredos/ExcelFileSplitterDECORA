using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class WorkbookPreparerTests
{
    private const string OriginalFormula = """
        let
            Zrodlo = Excel.Workbook(File.Contents("C:\dane\klienci.xlsx")),
            Naglowki = Table.PromoteHeaders(Zrodlo)
        in
            Naglowki
        """;

    private static SplitConfig Config() => new()
    {
        SplitColumn = "Nazwa Opiekuna Klienta",
        PowerPivot = new PowerPivotOptions { SplitTable = "dim_Klienci" },
    };

    private static QuerySpec Dimension() => new()
    {
        Name = "dim_Klienci",
        FilterColumn = "Nazwa Opiekuna Klienta",
        Cache = true,
    };

    [Fact]
    public void Wrapper_niesie_naglowek_z_numerem_wersji()
    {
        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), Dimension(), OriginalFormula);

        Assert.StartsWith("// [ExcelFileSplitter " + WorkbookPreparer.Version + "]", formula);
        Assert.True(WorkbookPreparer.IsPreparedByCurrentVersion(formula));
    }

    [Fact]
    public void Oryginalna_formula_zostaje_nietknieta_w_srodku()
    {
        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), Dimension(), OriginalFormula);

        foreach (string line in OriginalFormula.Split('\n'))
            Assert.Contains(line.Trim(), formula);
    }

    [Fact]
    public void Zapytanie_z_cache_dostaje_przelacznik_na_fnCache()
    {
        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), Dimension(), OriginalFormula);

        Assert.Contains("fnParam(\"UseCache\")", formula);
        Assert.Contains("fnCache(\"dim_Klienci\")", formula);
    }

    [Fact]
    public void Zapytanie_bez_cache_czyta_wylacznie_ze_zrodla()
    {
        var query = Dimension();
        query.Cache = false;

        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), query, OriginalFormula);

        Assert.DoesNotContain("fnCache", formula);
        Assert.Contains("__source = (", formula);
    }

    [Fact]
    public void Filtr_wymiaru_obsluguje_kilka_wartosci_naraz()
    {
        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), Dimension(), OriginalFormula);

        Assert.Contains("Text.Split(", formula);
        Assert.Contains("List.Contains(__values", formula);
        Assert.Contains("Record.Field(_, \"Nazwa Opiekuna Klienta\")", formula);
    }

    [Fact]
    public void Bez_filtra_wartosc_calkowita_przechodzi_bez_zmian()
    {
        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), Dimension(), OriginalFormula);

        Assert.Contains("if __splitValue = \"__ALL__\" then __source", formula);
    }

    [Fact]
    public void Fakty_sa_zawezane_zlaczeniem_z_przefiltrowanym_wymiarem()
    {
        var facts = new QuerySpec
        {
            Name = "fact_Zamowienia",
            FilterColumn = "Odbiorca materialow",
            KeyFrom = new KeyFrom { Query = "dim_Klienci", Column = "Identyfikator" },
            Cache = true,
        };

        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), facts, OriginalFormula);

        Assert.Contains("Table.Join(__source", formula);
        Assert.Contains("JoinKind.Inner", formula);
        Assert.Contains("Table.SelectColumns(dim_Klienci, {\"Identyfikator\"})", formula);
        Assert.Contains("Table.RemoveColumns(__joined, {\"__efs_key\"})", formula);
        Assert.DoesNotContain("List.Contains", formula);
    }

    [Fact]
    public void Nazwa_zapytania_ze_spacja_jest_cytowana_jako_identyfikator_M()
    {
        var facts = new QuerySpec
        {
            Name = "fact_Zamowienia",
            FilterColumn = "Odbiorca materialow",
            KeyFrom = new KeyFrom { Query = "dim Klienci V2", Column = "Identyfikator" },
        };

        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), facts, OriginalFormula);

        Assert.Contains("#\"dim Klienci V2\"", formula);
    }

    [Fact]
    public void Zapytanie_bez_kolumny_filtra_nie_jest_filtrowane()
    {
        var products = new QuerySpec { Name = "dim_Produkty", FilterColumn = null, Cache = true };

        string formula = WorkbookPreparer.BuildWrappedFormula(Config(), products, OriginalFormula);

        Assert.DoesNotContain("Table.SelectRows", formula);
        Assert.DoesNotContain("Table.Join", formula);
        Assert.Contains("__result =", formula);
    }
}
