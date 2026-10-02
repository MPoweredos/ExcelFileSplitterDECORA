using System.Windows;
using System.Windows.Threading;

namespace ExcelFileSplitter.Ui;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += Failed;
        base.OnStartup(e);
    }

    private void Failed(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        e.Handled = true;

        MessageBox.Show(
            $"Program napotkał błąd, którego nie przewidzieliśmy:\n\n{e.Exception.Message}\n\n" +
            "Okno zostanie zamknięte, a Excel uruchomiony przez program - wyłączony. " +
            "Pliki, które zdążyły przejść kontrole, zostają w folderze wynikowym.",
            "Podział plików Excela - błąd",
            MessageBoxButton.OK,
            MessageBoxImage.Error);

        try
        {
            MainWindow?.Close();
        }
        catch (InvalidOperationException)
        {
            Shutdown();
        }
    }
}
