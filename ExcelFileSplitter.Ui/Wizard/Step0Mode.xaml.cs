using System.Threading.Tasks;
using System.Windows.Controls;
using ExcelFileSplitter.Core;

namespace ExcelFileSplitter.Ui.Wizard;

public partial class Step0Mode : UserControl, IWizardStep
{
    private readonly IWizardHost _host;

    public Step0Mode(IWizardHost host)
    {
        _host = host;
        InitializeComponent();
    }

    public string Title => "Rodzaj pliku";

    public Task EnterAsync()
    {
        bool pivot = _host.State.Config.Mode == SplitMode.PowerPivot;
        TrybModelu.IsChecked = pivot;
        TrybZwykly.IsChecked = !pivot;
        return Task.CompletedTask;
    }

    public void Flush() => _host.State.Config.Mode = Chosen();

    public Task<string?> LeaveAsync()
    {
        _host.StartMode(Chosen());
        return Task.FromResult<string?>(null);
    }

    private SplitMode Chosen() => TrybModelu.IsChecked == true ? SplitMode.PowerPivot : SplitMode.Worksheet;
}
