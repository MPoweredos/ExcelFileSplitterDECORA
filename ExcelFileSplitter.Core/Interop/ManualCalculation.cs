using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileSplitter.Interop;

public sealed class ManualCalculation : IDisposable
{
    private readonly Excel.Application _app;
    private readonly Excel.XlCalculation _previous;
    private bool _disposed;

    public ManualCalculation(Excel.Application app)
    {
        _app = app;

        _previous = Com.Or(() => app.Calculation, Excel.XlCalculation.xlCalculationAutomatic);

        Com.Try(() => app.Calculation = Excel.XlCalculation.xlCalculationManual);
    }

    public void CalculateOnce() => Com.Try(() => _app.Calculate());

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Com.WhenClosing(() => _app.Calculation = _previous);
    }
}
