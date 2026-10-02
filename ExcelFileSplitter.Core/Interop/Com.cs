using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelFileSplitter.Interop;

public static class Com
{
    public static bool SaysNo(Exception exception) =>
        exception is COMException or InvalidCastException or MissingMemberException
                  or NotSupportedException or RuntimeBinderException;

    private static readonly int[] DeadServerCodes =
    [
        unchecked((int)0x800706BA),
        unchecked((int)0x800706BE),
        unchecked((int)0x800706BF),
        unchecked((int)0x80010108),
        unchecked((int)0x80010105),
        unchecked((int)0x800401FD),
    ];

    public static bool SessionIsDead(Exception exception) => exception switch
    {
        COMException com => Array.IndexOf(DeadServerCodes, com.HResult) >= 0,

        InvalidComObjectException => true,

        _ => false,
    };

    public static bool WorthRetryingOnFreshExcel(Exception exception) => exception switch
    {
        ObjectDisposedException or OperationCanceledException => false,

        _ when SaysNo(exception) || SessionIsDead(exception) => true,

        { InnerException: { } cause } => WorthRetryingOnFreshExcel(cause),

        _ => false,
    };

    public static bool Try(Action step)
    {
        try
        {
            step();
            return true;
        }
        catch (Exception exception) when (SaysNo(exception) && !SessionIsDead(exception))
        {
            return false;
        }
    }

    public static bool WhenClosing(Action step)
    {
        try
        {
            step();
            return true;
        }
        catch (Exception exception) when (SaysNo(exception) || SessionIsDead(exception))
        {
            return false;
        }
    }

    public static T Or<T>(Func<T> read, T fallback)
    {
        try
        {
            return read();
        }
        catch (Exception exception) when (SaysNo(exception) && !SessionIsDead(exception))
        {
            return fallback;
        }
    }
}
