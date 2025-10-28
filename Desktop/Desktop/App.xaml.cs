using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using Windows.ApplicationModel.Activation;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop;

public partial class App : Application
{
    public App()
    {
        InitializeComponent();

        ConfigureDI();
    }

    private static void ConfigureDI()
    {
        Ioc.Default.ConfigureServices(
            new ServiceCollection()
            .AddSingleton<MainViewModel>()
            .AddSingleton<IProjectFileService, ProjectFileService>()
            .AddSingleton<IDialogService, DialogService>()
            .BuildServiceProvider());
    }

    protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        IStorageFile? projectFile = null;

        var activationArgs = AppInstance.GetCurrent().GetActivatedEventArgs();
        if (activationArgs?.Kind == ExtendedActivationKind.File &&
            activationArgs.Data is IFileActivatedEventArgs fileArgs &&
            fileArgs.Files.Count > 0 &&
            fileArgs.Files[0] is IStorageFile file)
        {
            projectFile = file;
        }

        var mainWindow = new MainWindow(projectFile);

        var dialogService = Ioc.Default.GetRequiredService<IDialogService>();
        dialogService.SetWindow(mainWindow);

        mainWindow.Activate();
    }
}