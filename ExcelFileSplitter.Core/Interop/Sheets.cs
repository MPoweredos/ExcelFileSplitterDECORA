using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed record SheetInfo(string Name, bool IsVisible);

public static class Sheets
{
    public static List<SheetInfo> List(Excel.Workbook workbook)
    {
        var sheets = new List<SheetInfo>();
        foreach (object item in workbook.Sheets)
        {
            SheetInfo? info = Describe(item);
            if (info is not null) sheets.Add(info);
        }
        return sheets;
    }

    public static int DeleteAllExcept(Excel.Workbook workbook, ICollection<string> keep)
    {
        var kept = new HashSet<string>(keep, StringComparer.OrdinalIgnoreCase);
        int deleted = 0;

        for (int i = workbook.Sheets.Count; i >= 1; i--)
        {
            object item = workbook.Sheets[i];
            SheetInfo? info = Describe(item);
            if (info is null || kept.Contains(info.Name)) continue;

            switch (item)
            {
                case Excel.Worksheet worksheet:
                    worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    worksheet.Delete();
                    deleted++;
                    break;
                case Excel.Chart chart:
                    chart.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    chart.Delete();
                    deleted++;
                    break;
            }
        }
        return deleted;
    }

    private static SheetInfo? Describe(object sheet) => sheet switch
    {
        Excel.Worksheet worksheet => new SheetInfo(worksheet.Name, IsVisible(worksheet.Visible)),
        Excel.Chart chart => new SheetInfo(chart.Name, IsVisible(chart.Visible)),
        _ => null,
    };

    private static bool IsVisible(Excel.XlSheetVisibility visibility) =>
        visibility == Excel.XlSheetVisibility.xlSheetVisible;
}
