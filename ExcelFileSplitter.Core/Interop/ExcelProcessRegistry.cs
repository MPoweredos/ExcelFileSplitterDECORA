using System.Diagnostics;
using System.Globalization;

namespace ExcelFileSplitter.Interop;

public sealed record ExcelProcessEntry(int ExcelPid, long ExcelStartTicks, int OwnerPid, long OwnerStartTicks)
{
    public string Format() =>
        string.Join(';', new[] { ExcelPid, ExcelStartTicks, OwnerPid, OwnerStartTicks }
            .Select(number => number.ToString(CultureInfo.InvariantCulture)));

    public static ExcelProcessEntry? Parse(string line)
    {
        string[] parts = line.Split(';');
        if (parts.Length != 4) return null;

        var numbers = new long[4];
        for (int i = 0; i < 4; i++)
            if (!long.TryParse(parts[i].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out numbers[i])) return null;

        if (numbers[0] is <= 0 or > int.MaxValue || numbers[2] is <= 0 or > int.MaxValue) return null;
        return new ExcelProcessEntry((int)numbers[0], numbers[1], (int)numbers[2], numbers[3]);
    }
}

public sealed record RunningProcess(string Name, long StartTicks);

public static class ExcelProcessRegistry
{
    private static readonly object Gate = new();

    public static string FilePath { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelFileSplitter", "procesy-excela.txt");

    public static (List<ExcelProcessEntry> Kill, List<ExcelProcessEntry> Keep) Decide(
        IEnumerable<ExcelProcessEntry> entries, Func<int, RunningProcess?> lookup, Func<ExcelProcessEntry, bool> shouldKill)
    {
        var kill = new List<ExcelProcessEntry>();
        var keep = new List<ExcelProcessEntry>();
        foreach (ExcelProcessEntry entry in entries.Distinct())
        {
            RunningProcess? excel = lookup(entry.ExcelPid);
            bool stillOurs = excel is not null
                             && excel.StartTicks == entry.ExcelStartTicks
                             && excel.Name.Equals("EXCEL", StringComparison.OrdinalIgnoreCase);
            if (!stillOurs) continue;

            if (shouldKill(entry)) kill.Add(entry);
            else keep.Add(entry);
        }
        return (kill, keep);
    }

    public static bool OwnerIsGone(ExcelProcessEntry entry, Func<int, RunningProcess?> lookup) =>
        lookup(entry.OwnerPid) is not RunningProcess owner || owner.StartTicks != entry.OwnerStartTicks;

    public static void Register(int excelPid)
    {
        if (excelPid <= 0 || Lookup(excelPid) is not RunningProcess excel || Current() is not RunningProcess owner) return;

        var entry = new ExcelProcessEntry(excelPid, excel.StartTicks, Environment.ProcessId, owner.StartTicks);
        Update(entries => entries.Add(entry));
    }

    public static void Unregister(int excelPid)
    {
        if (excelPid <= 0) return;
        Update(entries => entries.RemoveAll(entry => entry.ExcelPid == excelPid && entry.OwnerPid == Environment.ProcessId));
    }

    public static int KillLeftByEarlierRuns() =>
        KillWhere(entry => OwnerIsGone(entry, Lookup));

    public static int KillOwn(int? except = null)
    {
        long? ownStart = Current()?.StartTicks;
        return KillWhere(entry => entry.OwnerPid == Environment.ProcessId
                                  && entry.OwnerStartTicks == ownStart
                                  && entry.ExcelPid != except);
    }

    private static int KillWhere(Func<ExcelProcessEntry, bool> shouldKill)
    {
        int killed = 0;
        Update(entries =>
        {
            (List<ExcelProcessEntry> kill, List<ExcelProcessEntry> keep) = Decide(entries, Lookup, shouldKill);
            entries.Clear();
            entries.AddRange(keep);
            foreach (ExcelProcessEntry entry in kill)
            {
                if (TryKill(entry.ExcelPid)) killed++;
                else entries.Add(entry);
            }
        });
        return killed;
    }

    public static bool TryKill(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (process.HasExited) return true;

            process.Kill(entireProcessTree: true);
            return process.WaitForExit(5_000);
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (Exception exception) when (exception is InvalidOperationException or System.ComponentModel.Win32Exception
                                              or NotSupportedException)
        {
            return false;
        }
    }

    private static RunningProcess? Current() => Lookup(Environment.ProcessId);

    private static RunningProcess? Lookup(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (process.HasExited) return null;
            return new RunningProcess(process.ProcessName, process.StartTime.ToUniversalTime().Ticks);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException
                                              or System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return null;
        }
    }

    private static void Update(Action<List<ExcelProcessEntry>> change)
    {
        lock (Gate)
        {
            try
            {
                using var mutex = new Mutex(false, @"Local\ExcelFileSplitter-procesy-excela");
                bool owned;
                try
                {
                    owned = mutex.WaitOne(TimeSpan.FromSeconds(5));
                }
                catch (AbandonedMutexException)
                {
                    owned = true;
                }
                if (!owned) return;

                try
                {
                    List<ExcelProcessEntry> entries = File.Exists(FilePath)
                        ? File.ReadAllLines(FilePath).Select(ExcelProcessEntry.Parse).OfType<ExcelProcessEntry>().ToList()
                        : [];
                    change(entries);

                    Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                    File.WriteAllLines(FilePath, entries.Distinct().Select(entry => entry.Format()));
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException
                                                  or System.Security.SecurityException or WaitHandleCannotBeOpenedException)
            {
            }
        }
    }
}
