﻿using AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Windows.AppLifecycle;
using Windows.ApplicationModel.Activation;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop;

public sealed partial class App
{
    public App()
    {
        InitializeComponent();

        ConfigureDependencyInjection();

        CreateDataStorage();
    }

    private static void ConfigureDependencyInjection()
    {
        Ioc.Default.ConfigureServices(
            new ServiceCollection()
                .AddSingleton<MainViewModel>()
                .AddSingleton<IProjectFileService, ProjectFileService>()
                .AddSingleton<IDialogService, DialogService>()
                .AddSingleton<IStorageService, StorageService>()
                .AddSingleton<IProductsRepository, ProductRepository>()
                .AddSingleton<IPenelopeClient, PenelopeClient>()
                .BuildServiceProvider());
    }

    private static void CreateDataStorage()
    {
        var storageService = Ioc.Default.GetRequiredService<IStorageService>();

        if (!storageService.DoesContainerExist("Credentials"))
            storageService.CreateContainer("Credentials");
    }

    protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        IStorageFile? projectFile = null;

        var activationArgs = AppInstance.GetCurrent().GetActivatedEventArgs();
        if (activationArgs is { Kind: ExtendedActivationKind.File, Data: IFileActivatedEventArgs fileArgs } &&
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