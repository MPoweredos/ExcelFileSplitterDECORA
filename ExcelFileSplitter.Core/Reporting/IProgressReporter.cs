namespace ExcelFileSplitter;

public interface IProgressReporter
{
    void Step(string message);

    void Detail(string message);

    void Warning(string message);
}

public sealed class NullProgressReporter : IProgressReporter
{
    public static readonly NullProgressReporter Instance = new();

    private NullProgressReporter() { }

    public void Step(string message) { }
    public void Detail(string message) { }
    public void Warning(string message) { }
}
