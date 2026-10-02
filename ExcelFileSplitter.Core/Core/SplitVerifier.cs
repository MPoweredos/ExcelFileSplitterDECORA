using System.Globalization;
using System.Text.RegularExpressions;
using ExcelFileSplitter.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Core;

public sealed record VerificationResult(
    string FilePath,
    IReadOnlyList<string> ExpectedValues,
    List<string> ValuesInModel,
    List<string> StaleSlicerItems,
    List<string> Leaks,
    List<string> NameCoincidences,
    List<string> SchemaDifferences,
    List<string> KeyProblems,
    long DimensionRowCount,
    string? Error) : IVerificationResult
{
    public List<string> StalePivotMembers { get; init; } = [];

    public List<string> ForeignValues => ValuesInModel.Where(value => !ExpectedValues.Contains(value, StringComparer.Ordinal)).ToList();

    public List<string> ValuesWithoutData => ExpectedValues.Where(value => !ValuesInModel.Contains(value, StringComparer.Ordinal)).ToList();

    public bool Ok => Error is null
                      && ValuesInModel.Count > 0
                      && ForeignValues.Count == 0
                      && StaleSlicerItems.Count == 0
                      && StalePivotMembers.Count == 0
                      && Leaks.Count == 0
                      && SchemaDifferences.Count == 0
                      && KeyProblems.Count == 0;

    public string Summary()
    {
        if (Error is not null) return $"BŁĄD WERYFIKACJI: {Error}";

        if (Ok)
        {
            string scope = ExpectedValues.Count == 1
                ? $"tylko \"{ExpectedValues[0]}\""
                : $"tylko wartości odbiorcy ({ValuesInModel.Count} z {ExpectedValues.Count})";
            var notes = new List<string>();
            if (ValuesWithoutData.Count > 0) notes.Add($"brak danych dla: {string.Join(", ", ValuesWithoutData)}");
            if (NameCoincidences.Count > 0) notes.Add($"zbiegi nazw z danymi odbiorcy: {string.Join(", ", NameCoincidences.Take(3))}");
            return $"OK: {scope}, {DimensionRowCount:N0} wierszy wymiaru; struktura modelu, klucze, fragmentatory, " +
                   "tabele przestawne i paczka czyste" +
                   (notes.Count > 0 ? $" ({string.Join("; ", notes)})" : "");
        }

        var problems = new List<string>();
        if (ValuesInModel.Count == 0) problems.Add("model jest pusty (filtr nie zwrócił nic?)");
        if (ForeignValues.Count > 0) problems.Add($"model zawiera obce wartości: {string.Join(", ", ForeignValues.Take(5))}");
        problems.AddRange(KeyProblems);
        if (SchemaDifferences.Count > 0)
            problems.Add($"struktura modelu inna niż w oryginale: {string.Join(" | ", SchemaDifferences.Take(5))}");
        if (StaleSlicerItems.Count > 0)
            problems.Add($"fragmentatory pokazują {StaleSlicerItems.Count} element(ów) spoza modelu: " +
                         string.Join(" | ", StaleSlicerItems.Take(4)));
        if (StalePivotMembers.Count > 0)
            problems.Add("tabele przestawne pamiętają pozycje spoza modelu odbiorcy, zapisane w pliku, choć ich " +
                         "nie widać: " + string.Join(" | ", StalePivotMembers.Take(4)));
        if (Leaks.Count > 0)
            problems.Add($"WYCIEK w paczce ({Leaks.Count}): " + string.Join(" | ", Leaks.Take(4)));
        return "WERYFIKACJA: " + string.Join("; ", problems);
    }
}

public static class SplitVerifier
{
    private static readonly Regex OlapSourcePattern = new(@"^\[(?<t>.+?)\]\.\[(?<c>.+?)\]$", RegexOptions.Compiled);

    public static VerificationResult VerifyRecipientFile(
        ExcelSession session, string path, SplitConfig config, Recipient recipient,
        IReadOnlyList<string> allSplitColumnValues, IReadOnlyList<ModelTableInfo> referenceTables,
        IProgressReporter reporter)
    {
        var expectedValues = recipient.Keys.Select(key => key.Values[0]).ToList();
        var forbiddenValues = allSplitColumnValues
            .Where(value => !expectedValues.Contains(value, StringComparer.OrdinalIgnoreCase))
            .ToList();

        Excel.Workbook? workbook = null;
        try
        {
            List<CachedModelField> cachedFields = PivotCacheMembers.Read(path);

            workbook = session.OpenWorkbook(path, readOnly: true);

            using var model = new ModelQuery(workbook);
            var valuesInModel = model.ReadDistinctValues(config.PowerPivot.SplitTable, config.SplitColumn);

            long dimensionRowCount = Com.Or(() => model.CountTableRows(config.PowerPivot.SplitTable), 0L);

            var tables = ModelSchema.ReadTables(workbook, model, countRows: false);
            var schemaDifferences = ModelSchema.FindDifferences(referenceTables, tables);
            var keyProblems = FindOrphanFactKeys(model, config);
            var staleSlicerItems = FindStaleSlicerItems(workbook, model, reporter);
            var stalePivotMembers = FindStalePivotMembers(cachedFields, model);

            HashSet<string>? legalInOtherColumns;
            try
            {
                legalInOtherColumns = model.FindValuesPresentInColumns(
                    tables.SelectMany(table => table.Columns.Select(column => (table.Name, column.Name))), forbiddenValues);
            }
            catch (Exception exception) when (Com.SaysNo(exception))
            {
                legalInOtherColumns = null;
            }

            ExcelSession.CloseWorkbook(workbook, save: false);
            workbook = null;

            (List<string> leaks, List<string> nameCoincidences) =
                ScanPackage(path, forbiddenValues, config, legalInOtherColumns);

            return new VerificationResult(
                path, expectedValues, valuesInModel, staleSlicerItems, leaks, nameCoincidences,
                schemaDifferences, keyProblems, dimensionRowCount, null)
            {
                StalePivotMembers = stalePivotMembers,
            };
        }
        catch (Exception ex)
        {
            return new VerificationResult(path, expectedValues, [], [], [], [], [], [], 0, ex.Message);
        }
        finally
        {
            ExcelSession.CloseWorkbook(workbook, save: false);
        }
    }

    private static (List<string> Leaks, List<string> NameCoincidences) ScanPackage(
        string path, IReadOnlyList<string> forbiddenValues, SplitConfig config, HashSet<string>? legalInOtherColumns)
    {
        var leaks = new List<string>();
        var nameCoincidences = new List<string>();

        foreach (ScanHit hit in PackageScanner.FindForbiddenValues(path, forbiddenValues, config.PowerPivot.SplitTable, [config.SplitColumn]))
        {
            if (!hit.IsQualifiedReference && legalInOtherColumns is not null && legalInOtherColumns.Contains(hit.Value))
            {
                if (!nameCoincidences.Contains(hit.Value, StringComparer.OrdinalIgnoreCase)) nameCoincidences.Add(hit.Value);
                continue;
            }

            string description = hit.Describe();
            if (!hit.IsQualifiedReference && legalInOtherColumns is null) description += " (nie udało się sprawdzić w modelu)";
            leaks.Add(description);
        }

        return (leaks, nameCoincidences);
    }

    private static List<string> FindOrphanFactKeys(ModelQuery model, SplitConfig config)
    {
        var problems = new List<string>();
        var keyFiltered = config.PowerPivot.Queries.Where(query => query.KeyFrom is not null && !string.IsNullOrWhiteSpace(query.FilterColumn));

        foreach (var query in keyFiltered)
        {
            try
            {
                long orphanCount = model.CountOrphanKeys(
                    query.Name, query.FilterColumn!, query.KeyFrom!.Query, query.KeyFrom.Column);
                if (orphanCount > 0)
                    problems.Add($"'{query.Name}' ma {orphanCount:N0} kluczy [{query.FilterColumn}] " +
                                 $"spoza wymiaru '{query.KeyFrom.Query}' - dane innych odbiorców");
            }
            catch (Exception ex)
            {
                problems.Add($"'{query.Name}': nie udało się sprawdzić kluczy ({ex.Message})");
            }
        }
        return problems;
    }

    private static List<string> FindStalePivotMembers(List<CachedModelField> fields, ModelQuery model)
    {
        var problems = new List<string>();
        var valuesByColumn = new Dictionary<string, ModelValues?>(StringComparer.OrdinalIgnoreCase);

        foreach (CachedModelField field in fields)
        {
            string columnKey = ModelQuery.ColumnKey(field.Table, field.Column);
            if (!valuesByColumn.TryGetValue(columnKey, out ModelValues? values))
            {
                values = Com.Or<ModelValues?>(
                    () => new ModelValues(model.ReadDistinctValues(field.Table, field.Column), CultureInfo.CurrentCulture),
                    null);
                valuesByColumn[columnKey] = values;
            }

            if (values is null)
            {
                problems.Add($"{field.Owner} [{field.Field}] -> nie da się sprawdzić, bo w modelu nie ma kolumny " +
                             $"'{field.Table}'[{field.Column}]");
                continue;
            }

            List<string> outside = PivotCacheMembers.OutsideModel(field, values);
            if (outside.Count > 0)
                problems.Add($"{field.Owner} [{field.Field}] -> {outside.Count}: {string.Join(", ", outside.Take(5))}" +
                             (outside.Count > 5 ? ", …" : ""));
        }
        return problems;
    }

    private static List<string> FindStaleSlicerItems(
        Excel.Workbook workbook, ModelQuery model, IProgressReporter reporter)
    {
        var staleItems = new List<string>();
        var valuesByColumn = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (SlicerInfo slicer in CachedData.ListSlicers(workbook, reporter))
        {
            Match match = OlapSourcePattern.Match(slicer.SourceName);
            if (!match.Success) continue;

            string table = match.Groups["t"].Value;
            string column = match.Groups["c"].Value;
            string cacheKey = table + "|" + column;

            if (!valuesByColumn.TryGetValue(cacheKey, out var allowedValues))
            {
                var fromModel = Com.Or<HashSet<string>?>(
                    () => new HashSet<string>(
                        model.ReadDistinctValues(table, column).Select(value => value.Trim()), StringComparer.OrdinalIgnoreCase),
                    null);
                if (fromModel is null) continue;
                allowedValues = fromModel;
                valuesByColumn[cacheKey] = allowedValues;
            }

            foreach (string caption in SlicerItems.OutsideData(slicer.Captions, allowedValues))
                staleItems.Add($"{slicer.Name}[{column}] -> {caption}");
        }
        return staleItems;
    }
}
