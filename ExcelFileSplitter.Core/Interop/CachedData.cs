using System.Text.RegularExpressions;
using ExcelFileSplitter.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed record SlicerInfo(string Name, string SourceName, List<string> Captions);

public static class CachedData
{
    private static readonly Regex OlapSourcePattern = new(@"^\[(?<t>.+?)\]\.\[(?<c>.+?)\]$", RegexOptions.Compiled);

    public static string FieldNameOf(string sourceName)
    {
        Match match = OlapSourcePattern.Match(sourceName);
        return match.Success ? match.Groups["c"].Value : sourceName.Trim();
    }

    public static PivotRefresh PurgePivotCaches(Excel.Workbook workbook, IProgressReporter reporter)
    {
        DateTime started = DateTime.Now;
        Dictionary<int, List<PivotTableRef>> tablesByCache = PivotTablesByCache(workbook);

        int refreshed = 0;
        int notLimited = 0;
        int olap = 0;
        var failures = new List<PivotRefreshFailure>();
        var failedDates = new HashSet<DateTime>();

        foreach (Excel.PivotCache pivotCache in workbook.PivotCaches())
        {
            if (Com.Or(() => pivotCache.OLAP, false)) olap++;
            else if (!Com.Try(() => pivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone))
                notLimited++;

            string? error = TryRefresh(pivotCache);
            if (error is null)
            {
                refreshed++;
                continue;
            }

            int index = Com.Or(() => pivotCache.Index, 0);
            List<PivotTableRef> tables = tablesByCache.TryGetValue(index, out List<PivotTableRef>? found) ? found : [];
            string? source = PivotSource.Describe(Com.Or<object?>(() => pivotCache.SourceData, null));
            failures.Add(new PivotRefreshFailure(tables, error, source));

            DateTime? refreshedAt = Com.Or<DateTime?>(() => pivotCache.RefreshDate, null);
            if (refreshedAt is not null) failedDates.Add(refreshedAt.Value);
        }

        if (olap > 0)
            reporter.Detail($"{olap} cache tabel przestawnych opiera się na modelu danych - pozycje biorą wprost " +
                            "z modelu, więc kasowanie nieaktualnych ich nie dotyczy");

        if (notLimited > 0)
            reporter.Warning($"{notLimited} cache tabel przestawnych nie przyjął ustawienia kasowania nieaktualnych " +
                             "pozycji - mogą w nich zostać wartości, których w danych już nie ma");

        return new PivotRefresh(started, refreshed, failures, failedDates);
    }

    public static void ReportRefreshFailures(Excel.Workbook workbook, PivotRefresh refresh, IProgressReporter reporter)
    {
        if (refresh.Failures.Count == 0) return;

        var sheetsInFile = new HashSet<string>(
            Sheets.List(workbook).Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);

        foreach (PivotRefreshFailure failure in refresh.Failures)
        {
            (bool isWarning, string message) = failure.Report(sheetsInFile);
            if (isWarning) reporter.Warning(message);
            else reporter.Detail(message);
        }
    }

    public static List<string> FindStalePivotTables(Excel.Workbook workbook, PivotRefresh refresh)
    {
        var stale = new List<string>();

        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            List<Excel.PivotTable>? tables = PivotTablesOn(sheet, out int unreadable);
            if ((tables is null || unreadable > 0) && refresh.Failures.Count > 0)
                stale.Add($"arkusz '{sheet.Name}' - nie da się odczytać wszystkich jego tabel przestawnych, " +
                          "a część cache'y się nie odświeżyła");

            foreach (Excel.PivotTable table in tables ?? [])
            {
                Excel.PivotCache? cache = Com.Or<Excel.PivotCache?>(() => table.PivotCache(), null);
                if (cache is not null && Com.Or(() => cache.OLAP, false)) continue;

                DateTime? refreshedAt = cache is null ? null : Com.Or<DateTime?>(() => cache.RefreshDate, null);
                if (!refresh.IsStale(refreshedAt)) continue;

                var name = new PivotTableRef(Com.Or(() => table.Name, "?"), sheet.Name);
                stale.Add(refreshedAt is null
                    ? $"{name} - nie da się ustalić, kiedy odświeżono jej cache"
                    : $"{name} - cache z {refreshedAt.Value:yyyy-MM-dd HH\\:mm\\:ss}");
            }
        }

        return stale;
    }

    private static string? TryRefresh(Excel.PivotCache pivotCache)
    {
        try
        {
            pivotCache.Refresh();
            return null;
        }
        catch (Exception exception) when (Com.SaysNo(exception) && !Com.SessionIsDead(exception))
        {
            return ExcelMessage.FirstSentence(exception.Message);
        }
    }

    private static Dictionary<int, List<PivotTableRef>> PivotTablesByCache(Excel.Workbook workbook)
    {
        var byCache = new Dictionary<int, List<PivotTableRef>>();

        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            foreach (Excel.PivotTable table in PivotTablesOn(sheet, out _) ?? [])
            {
                int cacheIndex = Com.Or(() => table.CacheIndex, 0);
                if (!byCache.TryGetValue(cacheIndex, out List<PivotTableRef>? tables)) byCache[cacheIndex] = tables = [];
                tables.Add(new PivotTableRef(Com.Or(() => table.Name, "?"), sheet.Name));
            }
        }

        return byCache;
    }

    private static List<Excel.PivotTable>? PivotTablesOn(Excel.Worksheet sheet, out int unreadable)
    {
        unreadable = 0;

        var collection = Com.Or<Excel.PivotTables?>(() => (Excel.PivotTables)sheet.PivotTables(), null);
        if (collection is null) return null;

        int? count = Com.Or<int?>(() => collection.Count, null);
        if (count is null) return null;

        var tables = new List<Excel.PivotTable>();
        for (int i = 1; i <= count; i++)
        {
            int index = i;
            var table = Com.Or<Excel.PivotTable?>(() => (Excel.PivotTable)collection.Item(index), null);
            if (table is null) unreadable++;
            else tables.Add(table);
        }

        return tables;
    }

    public static int RefreshChartCaches(Excel.Workbook workbook, IProgressReporter reporter)
    {
        int refreshed = 0;
        int failed = 0;

        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            var chartObjects = Com.Or<Excel.ChartObjects?>(() => (Excel.ChartObjects)sheet.ChartObjects(), null);
            if (chartObjects is null) continue;

            for (int i = 1; i <= chartObjects.Count; i++)
            {
                int index = i;
                if (Com.Try(() => ((Excel.ChartObject)chartObjects.Item(index)).Chart.Refresh())) refreshed++;
                else failed++;
            }
        }

        foreach (Excel.Chart chart in Com.Or<IEnumerable<Excel.Chart>>(() => workbook.Charts.Cast<Excel.Chart>().ToList(), []))
        {
            if (Com.Try(() => chart.Refresh())) refreshed++;
            else failed++;
        }

        if (failed > 0)
            reporter.Warning($"{failed} wykres(ów) nie dało się odświeżyć - mogą pokazywać dane sprzed zawężenia");

        return refreshed;
    }

    public static void DisableRefreshOnOpen(Excel.Workbook workbook, IProgressReporter reporter)
    {
        int connectionsRefused = 0;

        foreach (Excel.WorkbookConnection connection in workbook.Connections)
        {
            if (!RefreshesOnOpen(connection)) continue;

            bool oledb = Com.Try(() => connection.OLEDBConnection.RefreshOnFileOpen = false);
            bool odbc = Com.Try(() => connection.ODBCConnection.RefreshOnFileOpen = false);
            if ((!oledb && !odbc) || RefreshesOnOpen(connection)) connectionsRefused++;
        }

        int cachesRefused = 0;
        foreach (Excel.PivotCache pivotCache in workbook.PivotCaches())
        {
            if (!Com.Or(() => pivotCache.RefreshOnFileOpen, false)) continue;

            bool disabled = Com.Try(() => pivotCache.RefreshOnFileOpen = false);
            if (!disabled || Com.Or(() => pivotCache.RefreshOnFileOpen, false)) cachesRefused++;
        }

        if (connectionsRefused > 0 || cachesRefused > 0)
            reporter.Warning($"nie udało się wyłączyć odświeżania przy otwarciu: połączenia {connectionsRefused}, " +
                             $"cache tabel przestawnych {cachesRefused} - plik może próbować sięgnąć do źródeł u odbiorcy");
    }

    private static bool RefreshesOnOpen(Excel.WorkbookConnection connection) =>
        Com.Or(() => connection.OLEDBConnection.RefreshOnFileOpen, false)
        || Com.Or(() => connection.ODBCConnection.RefreshOnFileOpen, false);

    public static int RemoveSlicersBoundToColumn(Excel.Workbook workbook, string column, IProgressReporter reporter)
    {
        int removed = 0;
        int failed = 0;

        Excel.SlicerCaches slicerCaches = workbook.SlicerCaches;
        for (int i = slicerCaches.Count; i >= 1; i--)
        {
            int index = i;
            var slicerCache = Com.Or<Excel.SlicerCache?>(() => (Excel.SlicerCache)slicerCaches[index], null);
            if (slicerCache is null)
            {
                failed++;
                continue;
            }

            string sourceName = Com.Or(() => slicerCache.SourceName ?? "", "");
            if (!sourceName.Contains(column, StringComparison.OrdinalIgnoreCase)) continue;

            if (Com.Try(() => slicerCache.Delete())) removed++;
            else failed++;
        }

        if (failed > 0)
            reporter.Warning($"{failed} fragmentator(ów) nie dało się usunąć - sprawdź je w pliku wynikowym");

        return removed;
    }

    public static List<SlicerInfo> ListSlicers(Excel.Workbook workbook, IProgressReporter reporter)
    {
        var slicers = new List<SlicerInfo>();

        var slicerCaches = Com.Or<Excel.SlicerCaches?>(() => workbook.SlicerCaches, null);
        if (slicerCaches is null) return slicers;

        int unreadable = 0;
        for (int i = 1; i <= slicerCaches.Count; i++)
        {
            int index = i;
            var slicerCache = Com.Or<Excel.SlicerCache?>(() => (Excel.SlicerCache)slicerCaches[index], null);
            if (slicerCache is null)
            {
                unreadable++;
                continue;
            }

            string sourceName = Com.Or<string?>(() => slicerCache.SourceName, null) ?? "";
            if (sourceName.Length == 0)
            {
                unreadable++;
                continue;
            }

            string name = Com.Or(() => slicerCache.Name, sourceName);
            slicers.Add(new SlicerInfo(name, sourceName, ReadItemCaptions(slicerCache)));
        }

        if (unreadable > 0)
            reporter.Warning($"{unreadable} fragmentator(ów) nie dało się odczytać - kontrola ich nie obejmie, " +
                             "obejrzyj je w pliku wynikowym");

        return slicers;
    }

    public static bool IsSyntheticMember(string caption) =>
        caption.Length == 0
        || caption is "(All)" or "(Wszystkie)" or "(blank)" or "(puste)" or "(Blank)" or "(Puste)";

    private static List<string> ReadItemCaptions(Excel.SlicerCache slicerCache)
    {
        var captions = new List<string>();

        Com.Try(() =>
        {
            foreach (Excel.SlicerCacheLevel level in slicerCache.SlicerCacheLevels)
                foreach (Excel.SlicerItem item in level.SlicerItems)
                    captions.Add(ReadCaption(item));
        });
        if (captions.Count > 0) return captions;

        Com.Try(() =>
        {
            foreach (Excel.SlicerItem item in slicerCache.SlicerItems) captions.Add(ReadCaption(item));
        });

        return captions;
    }

    private static string ReadCaption(Excel.SlicerItem item) =>
        Com.Or<string?>(() => item.Caption, null)
        ?? Com.Or<string?>(() => item.Name, null)
        ?? "";
}
