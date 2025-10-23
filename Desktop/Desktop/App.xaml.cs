using AdactaInternational.AdactaReportsShoppingBag.Desktop.Project;
using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using Windows.ApplicationModel.Activation;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop;

public partial class App : Application
{
    private MainWindow? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        var activationArgs = AppInstance.GetCurrent().GetActivatedEventArgs();

        _window = new MainWindow();
        _window.Activate();

        if (activationArgs.Kind == ExtendedActivationKind.File && activationArgs.Data is IFileActivatedEventArgs fileArgs && fileArgs.Files[0] is IStorageFile file)
        {

            if (ProjectManager.IsProjectFileValid(file, out ReportPrj? reportPrj))
            {
                _window.MainViewModel.ReportPrj = reportPrj;
                _window.MainViewModel.IsLoaded = true;
            }
            else
            {
                _window.MainViewModel.IsLoaded = false;
            }
        }
    }
}