namespace ExcelFileSplitter.Core;

public static class RefreshClock
{
    public static DateTime Threshold(DateTime refreshStarted) =>
        refreshStarted
            .AddTicks(-(refreshStarted.Ticks % TimeSpan.TicksPerSecond))
            .AddSeconds(-1);
}
