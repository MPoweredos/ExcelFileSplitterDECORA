using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class ModelSchemaTests
{
    private const int Tekst = 130;
    private const int Data = 7;

    private static ModelTableInfo Table(string name, params (string Name, int DataType)[] columns) =>
        new(name, columns.Select(c => new ModelColumn(c.Name, c.DataType, false)).ToList(), 0);

    [Fact]
    public void Identyczne_modele_nie_maja_roznic()
    {
        var expected = new[] { Table("dim_Klienci", ("Opiekun", Tekst), ("Identyfikator", Tekst)) };
        var actual = new[] { Table("dim_Klienci", ("Opiekun", Tekst), ("Identyfikator", Tekst)) };

        Assert.Empty(ModelSchema.FindDifferences(expected, actual));
    }

    [Fact]
    public void Kolumna_dolozona_przez_PowerPivot_jest_wykrywana()
    {
        var expected = new[] { Table("dim_Klienci", ("Kohorta", Tekst)) };
        var actual = new[] { Table("dim_Klienci", ("Kohorta", Tekst), ("Kohorta1", Tekst)) };

        var differences = ModelSchema.FindDifferences(expected, actual);

        Assert.Single(differences);
        Assert.Contains("Kohorta1", differences[0]);
        Assert.Contains("dodatkowa kolumna", differences[0]);
    }

    [Fact]
    public void Zgubiona_kolumna_jest_wykrywana()
    {
        var expected = new[] { Table("dim_Klienci", ("Opiekun", Tekst), ("Identyfikator", Tekst)) };
        var actual = new[] { Table("dim_Klienci", ("Opiekun", Tekst)) };

        var differences = ModelSchema.FindDifferences(expected, actual);

        Assert.Single(differences);
        Assert.Contains("Identyfikator", differences[0]);
    }

    [Fact]
    public void Zmieniony_typ_kolumny_jest_wykrywany()
    {
        var expected = new[] { Table("fact_Zamowienia", ("Data", Data)) };
        var actual = new[] { Table("fact_Zamowienia", ("Data", Tekst)) };

        var differences = ModelSchema.FindDifferences(expected, actual);

        Assert.Single(differences);
        Assert.Contains("Data", differences[0]);
        Assert.Contains("tekst", differences[0]);
    }

    [Fact]
    public void Brakujaca_i_dodatkowa_tabela_sa_wykrywane()
    {
        var expected = new[] { Table("dim_Klienci", ("Opiekun", Tekst)) };
        var actual = new[] { Table("dim_Produkty", ("Nazwa", Tekst)) };

        var differences = ModelSchema.FindDifferences(expected, actual);

        Assert.Equal(2, differences.Count);
        Assert.Contains(differences, d => d.Contains("brak tabeli") && d.Contains("dim_Klienci"));
        Assert.Contains(differences, d => d.Contains("dodatkowa tabela") && d.Contains("dim_Produkty"));
    }

    [Fact]
    public void Wielkosc_liter_w_nazwach_nie_jest_roznica()
    {
        var expected = new[] { Table("dim_Klienci", ("Opiekun", Tekst)) };
        var actual = new[] { Table("DIM_KLIENCI", ("OPIEKUN", Tekst)) };

        Assert.Empty(ModelSchema.FindDifferences(expected, actual));
    }
}
