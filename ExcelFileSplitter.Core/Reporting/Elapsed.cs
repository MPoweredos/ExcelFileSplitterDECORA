namespace ExcelFileSplitter;

public static class Elapsed
{
    public static string Text(this TimeSpan elapsed) =>
        elapsed.TotalSeconds < 10
            ? $"{elapsed.TotalSeconds:0.0}s"
            : $"{elapsed:mm\\:ss}";
}
