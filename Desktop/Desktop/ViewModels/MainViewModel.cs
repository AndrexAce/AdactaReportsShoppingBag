using AdactaInternational.AdactaReportsShoppingBag.Desktop.Exceptions;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class MainViewModel(IProjectFileService projectFileService, IDialogService dialogService, IStorageService storageService, IProductsRepository productsRepository)
    : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaveStateText), nameof(SaveButtonVisibility))]
    public partial bool? IsProjectEdited { get; private set; } = null;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsProjectEdited), nameof(NavigationViewMenuItems))]
    public partial ReportPrj? ReportProject { get; private set; } = null;

    public string SaveStateText => IsProjectEdited switch
    {
        true => "– Modifiche non salvate",
        false => "– Nessuna modifica non salvata",
        _ => string.Empty
    };

    public Visibility SaveButtonVisibility =>
        IsProjectEdited is null or false ? Visibility.Collapsed : Visibility.Visible;

    public ObservableCollection<Product>? NavigationViewMenuItems => new(ReportProject?.Products ?? []);

    private string? _projectFilePath;

    [RelayCommand]
    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    private async Task NewProjectAsync()
    {
        var (newProjectChoice, projectCode, projectName) =
            await dialogService.ShowNewProjectDialogAsync("Crea nuovo progetto", "Crea", "Annulla");

        if (newProjectChoice is ContentDialogResult.None) return;

        IEnumerable<Product>? products = null;

        while (products is null)
        {
            try
            {
                products = await productsRepository.GetProductsAsync(projectCode);
            }
            catch (PenelopeNotFoundException)
            {
                await dialogService.ShowInformationDialogAsync("Codice progetto non valido", "Il progetto inserito è errato o non esistente.", "Ok");

                return;
            }
            catch (PenelopeAuthenticationException)
            {
                var (credentialsChoice, penelopeUsername, penelopePassword) =
                    await dialogService.ShowPenelopeCredentialsDialogAsync("Account Penelope", "Conferma", "Annulla");

                if (credentialsChoice is ContentDialogResult.None) return;

                storageService.SaveData("Credentials", "Username", penelopeUsername);
                storageService.SaveData("Credentials", "Password", penelopePassword);
            }
        }

        var userChosenFolder = await dialogService.ShowFolderPicker();

        if (userChosenFolder is null) return;

        var project = new ReportPrj
        {
            ProjectCode = projectCode,
            ProjectName = projectName,
            Version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(),
            Products = products
        };

        _projectFilePath = projectFileService.CreateProjectFolder(project, userChosenFolder.Path);

        if (_projectFilePath is null) return;

        ReportProject = project;
        IsProjectEdited = false;
    }

    [RelayCommand]
    private async Task OpenProjectAsync()
    {
        var file = await dialogService.ShowFileOpenPickerAsync();

        if (file == null) return;

        ReportProject = await projectFileService.LoadProjectFileAsync(file);

        if (ReportProject is null)
        {
            await dialogService.ShowInformationDialogAsync("Progetto non caricato",
                "Il file del progetto è danneggiato.", "Ok");
        }
        else
        {
            _projectFilePath = file.Path;
            IsProjectEdited = false;
        }
    }

    [RelayCommand]
    private async Task SaveProjectAsync()
    {
        if (ReportProject == null || _projectFilePath == null) return;

        await projectFileService.SaveProjectFileAsync(ReportProject, _projectFilePath);

        IsProjectEdited = false;
    }

    [RelayCommand]
    private Task OpenInfoAsync()
    {
        return dialogService.ShowCreditsDialogAsync();
    }

    public async Task LoadProjectFileAsync(IStorageFile file)
    {
        ReportProject = await projectFileService.LoadProjectFileAsync(file);

        if (ReportProject is null)
        {
            await dialogService.ShowInformationDialogAsync("Progetto non caricato",
                "Il file del progetto è danneggiato.", "Ok");
        }
        else
        {
            _projectFilePath = file.Path;
            IsProjectEdited = false;
        }
    }
}