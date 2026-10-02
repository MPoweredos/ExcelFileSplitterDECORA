namespace ExcelFileSplitter.Core;

public sealed partial class SplitConfig
{
    public void Validate()
    {
        if (Passes.Count > 0)
        {
            ValidatePasses();
            return;
        }

        ValidateSinglePass();
    }

    private void ValidatePasses()
    {
        if (!string.IsNullOrWhiteSpace(SplitColumn) || SplitColumns.Count > 0
            || Values.Count > 0 || Keys.Count > 0 || Groups.Count > 0)
            throw new InvalidOperationException(
                "Ustawiona jest sekcja Passes i jednocześnie pola pojedynczego podziału " +
                "(SplitColumn, SplitColumns, Values, Keys albo Groups) - zostaw jedno z nich. " +
                "Przy kilku podziałach każdy opisuje swoje kolumny i wartości u siebie.");

        if (Mode != SplitMode.Worksheet)
            throw new InvalidOperationException(
                "Kilka niezależnych podziałów działa wyłącznie w trybie Worksheet (zwykły plik). " +
                "W trybie PowerPivot filtr siedzi w kodzie M, a przygotowanie skoroszytu dotyczy " +
                "jednej kolumny - żeby podzielić po drugiej, potrzebny jest osobny przebieg.");

        foreach (SplitPass pass in Passes)
        {
            if (string.IsNullOrWhiteSpace(pass.Name))
                throw new InvalidOperationException(
                    "Każdy podział w sekcji Passes musi mieć Name - trafia ono do nazwy pliku " +
                    "(znacznik {podzial}) i do logu.");
        }

        var duplicate = Passes
            .GroupBy(pass => pass.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);

        if (duplicate is not null)
            throw new InvalidOperationException(
                $"Podział o nazwie '{duplicate.Key}' występuje więcej niż raz - nazwy muszą się różnić, " +
                "bo po nich rozpoznaje się pliki.");

        foreach (SplitPass pass in Passes) ForPass(pass).Validate();
    }

    private void ValidateSinglePass()
    {
        RequireValue(SourceWorkbook, nameof(SourceWorkbook));
        RequireValue(OutputFolder, nameof(OutputFolder));
        ValidateSplitColumns();

        if (!File.Exists(SourceWorkbook))
            throw new FileNotFoundException($"Nie znaleziono skoroszytu źródłowego: {SourceWorkbook}");

        if (!ExcelFileTypes.IsWorkbook(SourceWorkbook))
            throw new InvalidOperationException(
                $"SourceWorkbook nie jest plikiem Excela: {SourceWorkbook}. " +
                $"Wskaż plik {ExcelFileTypes.Listed}.");

        if (Mode != SplitMode.Worksheet && !ExcelFileTypes.CanHoldDataModel(SourceWorkbook))
            throw new InvalidOperationException(
                $"Plik {Path.GetFileName(SourceWorkbook)} nie może trzymać modelu danych - " +
                "tryb PowerPivot wymaga xlsx, xlsm albo xlsb.");

        if (SheetsToKeep.Count == 0 && !Groups.Any(group => group.SheetsToKeep.Count > 0))
            throw new InvalidOperationException(
                "SheetsToKeep jest puste i żadna grupa nie podaje własnych arkuszy - plik wynikowy nie może nie mieć arkuszy.");

        if (!FileNameTemplate.Contains("{value}") && !FileNameTemplate.Contains("{name}"))
            throw new InvalidOperationException("FileNameTemplate musi zawierać {name} albo {value}.");

        if (Groups.Count > 0 && (Values.Count > 0 || Keys.Count > 0))
            throw new InvalidOperationException(
                "Ustawione są jednocześnie Groups i lista pojedynczych plików (Values albo Keys) - zostaw jedno z nich.");

        ValidateKeyArity();

        if (Mode == SplitMode.PowerPivot) ValidatePowerPivot();
        else ValidateWorksheet();

        static void RequireValue(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException($"Brak wymaganego pola konfiguracji: {fieldName}");
        }
    }

    private void ValidateSplitColumns()
    {
        if (SplitColumns.Count > 0 && !string.IsNullOrWhiteSpace(SplitColumn))
            throw new InvalidOperationException(
                $"Ustawione są jednocześnie SplitColumn ('{SplitColumn}') i SplitColumns " +
                $"({string.Join(", ", SplitColumns)}) - zostaw jedno z nich.");

        IReadOnlyList<string> columns = SplitColumnNames;

        if (columns.Count == 0)
            throw new InvalidOperationException(
                "Brak wymaganego pola konfiguracji: SplitColumn (albo SplitColumns przy podziale po kilku kolumnach)");

        var duplicate = columns
            .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);

        if (duplicate is not null)
            throw new InvalidOperationException(
                $"Kolumna '{duplicate.Key}' występuje w SplitColumns więcej niż raz - każda kolumna podziału może być tylko jedna.");
    }

    private void ValidateKeyArity()
    {
        int columns = SplitColumnNames.Count;

        if (columns > 1)
        {
            if (Values.Count > 0)
                throw new InvalidOperationException(
                    $"Przy podziale po {columns} kolumnach Values nie wystarczy - każdy plik opisuje kombinacja wartości. " +
                    "Użyj Keys, np. \"Keys\": [[\"Północ\", \"Kowalski\"], [\"Południe\", \"Nowak\"]].");

            RecipientGroup? withValues = Groups.FirstOrDefault(group => group.Values.Count > 0);
            if (withValues is not null)
                throw new InvalidOperationException(
                    $"Grupa '{withValues.Name}' używa Values, a podział idzie po {columns} kolumnach - użyj w niej Keys.");
        }

        CheckArity(this.RequestedKeys(), "Keys");
        foreach (RecipientGroup group in Groups) CheckArity(group.RequestedKeys(), $"grupa '{group.Name}'");

        void CheckArity(IReadOnlyList<SplitKey> keys, string where)
        {
            foreach (SplitKey key in keys)
            {
                if (key.Count != columns)
                    throw new InvalidOperationException(
                        $"{where}: kombinacja [{key.Display}] ma {key.Count} wartości, a kolumn podziału jest {columns} " +
                        $"({string.Join(", ", SplitColumnNames)}).");
            }
        }
    }

    private void ValidatePowerPivot()
    {
        if (SplitColumns.Count > 0)
            throw new InvalidOperationException(
                "Podział po kilku kolumnach działa wyłącznie w trybie Worksheet (zwykły plik). " +
                "W trybie PowerPivot filtr siedzi w kodzie M i przyjmuje jedną kolumnę - użyj SplitColumn.");

        if (string.IsNullOrWhiteSpace(PowerPivot.SplitTable))
            throw new InvalidOperationException("Brak wymaganego pola konfiguracji: PowerPivot.SplitTable");

        if (PowerPivot.UseCache && string.IsNullOrWhiteSpace(PowerPivot.CacheFolder))
            throw new InvalidOperationException("PowerPivot.UseCache = true wymaga ustawionego PowerPivot.CacheFolder.");

        foreach (var query in PowerPivot.Queries)
        {
            if (string.Equals(query.Name, PowerPivot.SplitTable, StringComparison.OrdinalIgnoreCase)
                && query.KeyFrom is null
                && !string.IsNullOrWhiteSpace(query.FilterColumn)
                && !string.Equals(query.FilterColumn, SplitColumn, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"Zapytanie '{query.Name}' filtruje po '{query.FilterColumn}', a SplitColumn to '{SplitColumn}'. " +
                    "Filtr i podział muszą dotyczyć tej samej kolumny.");
            }
        }
    }

    private void ValidateWorksheet()
    {
        if (string.IsNullOrWhiteSpace(Worksheet.Table) && string.IsNullOrWhiteSpace(Worksheet.Sheet))
            throw new InvalidOperationException(
                "Tryb Worksheet wymaga wskazania danych: Worksheet.Table (nazwa tabeli Excela) " +
                "albo Worksheet.Sheet (arkusz z danymi).");

        if (Worksheet.HeaderRow < 1)
            throw new InvalidOperationException(
                $"Worksheet.HeaderRow musi być dodatni - jest {Worksheet.HeaderRow}.");

        if (PowerPivot.UseCache)
            throw new InvalidOperationException("PowerPivot.UseCache dotyczy wyłącznie trybu PowerPivot - w trybie Worksheet nie ma modelu do zrzucenia.");

        if (PowerPivot.Queries.Count > 0)
            throw new InvalidOperationException(
                "PowerPivot.Queries opisują zapytania Power Query i dotyczą wyłącznie trybu PowerPivot. " +
                "Jeśli plik naprawdę ma zapytania, to nie jest zwykły plik - użyj trybu PowerPivot.");
    }
}
