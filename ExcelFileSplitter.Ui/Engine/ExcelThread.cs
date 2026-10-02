using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Engine;

public sealed class ExcelThread : IDisposable
{
    private readonly BlockingCollection<Action> _jobs = new();
    private readonly Thread _thread;
    private SplitWorkspace? _workspace;
    private bool _disposed;

    public ExcelThread()
    {
        _thread = new Thread(Loop, 16 * 1024 * 1024)
        {
            IsBackground = true,
            Name = "Excel (STA)",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> Run<T>(Func<SplitWorkspace, T> job)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        Enqueue(() =>
        {
            try { completion.SetResult(job(_workspace!)); }
            catch (Exception ex) { completion.SetException(ex); }
        });
        return completion.Task;
    }


    private void Enqueue(Action job)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _jobs.Add(job);
    }

    private void Loop()
    {
        _workspace = new SplitWorkspace();
        try
        {
            foreach (Action job in _jobs.GetConsumingEnumerable()) job();
        }
        finally
        {
            _workspace.Dispose();
            _workspace = null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _jobs.CompleteAdding();

        if (_thread.Join(TimeSpan.FromSeconds(30))) _jobs.Dispose();
    }
}
