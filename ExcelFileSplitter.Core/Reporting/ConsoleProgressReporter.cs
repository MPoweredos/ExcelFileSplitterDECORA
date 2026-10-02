namespace ExcelFileSplitter;

public sealed class ConsoleProgressReporter(string command) : IProgressReporter
{
    public void Step(string message) => Console.WriteLine($"[{command}] {message}");

    public void Detail(string message) => Console.WriteLine($"         {message}");

    public void Warning(string message) => Console.WriteLine($"[{command}] uwaga: {message}");
}
