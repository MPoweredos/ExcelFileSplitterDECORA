using System.Globalization;
using ExcelFileSplitter.Interop;

namespace ExcelFileSplitter.Core;

public static class SlicerItems
{
    public static List<string> OutsideData(
        IEnumerable<string> captions, IReadOnlySet<string> allowedValues, Func<string, bool>? alsoIgnore = null)
    {
        var outside = new List<string>();
        foreach (string raw in captions)
        {
            string caption = raw.Trim();
            if (CachedData.IsSyntheticMember(caption)) continue;
            if (alsoIgnore is not null && alsoIgnore(caption)) continue;
            if (!allowedValues.Contains(caption)) outside.Add(caption);
        }
        return outside;
    }

    public static bool LooksLikeNumberOrDate(string caption)
    {
        if (!caption.Any(char.IsDigit)) return false;
        if (double.TryParse(caption, NumberStyles.Any, CultureInfo.CurrentCulture, out _)) return true;
        if (double.TryParse(caption, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) return true;
        return DateTime.TryParse(caption, CultureInfo.CurrentCulture, DateTimeStyles.None, out _);
    }
}
