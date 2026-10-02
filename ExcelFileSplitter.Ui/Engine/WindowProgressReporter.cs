using System.Windows.Threading;

namespace ExcelFileSplitter.Ui.Engine;

public enum ProgressKind
{
    Step,

    Detail,

    Warning,
}

public sealed record ProgressLine(ProgressKind Kind, string Text)
{
    public string Display => Kind switch
    {
        ProgressKind.Detail => "      " + Text,
        ProgressKind.Warning => "uwaga: " + Text,
        _ => Text,
    };
}

public sealed class WindowProgressReporter(Dispatcher dispatcher, Action<ProgressLine> show) : IProgressReporter
{
    public void Step(string message) => Post(new ProgressLine(ProgressKind.Step, message));

    public void Detail(string message) => Post(new ProgressLine(ProgressKind.Detail, message));

    public void Warning(string message) => Post(new ProgressLine(ProgressKind.Warning, message));

    private void Post(ProgressLine line) => dispatcher.InvokeAsync(() => show(line));
}
