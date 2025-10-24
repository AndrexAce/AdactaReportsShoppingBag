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
            _window.MainViewModel.LoadProjectFile(file);
        }
    }
}