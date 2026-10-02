using System.Collections.ObjectModel;
using System.Threading.Tasks;
using ExcelFileSplitter.Core;
using ExcelFileSplitter.Ui.Engine;

namespace ExcelFileSplitter.Ui.Wizard;

public interface IWizardStep
{
    string Title { get; }

    Task EnterAsync();

    void Flush();

    Task<string?> LeaveAsync();
}

public interface IWizardHost
{
    WizardState State { get; }

    Task<T> RunAsync<T>(string what, Func<SplitWorkspace, IProgressReporter, T> job, bool lockContent = true);

    Task<T> RunCancellableAsync<T>(
        string what, Func<SplitWorkspace, IProgressReporter, CancellationToken, T> job);

    void Cancel();

    void Log(ProgressLine line);

    ObservableCollection<ProgressLine> LogLines { get; }

    void StartMode(SplitMode mode);

    void GoToStep<TStep>() where TStep : IWizardStep;
}
