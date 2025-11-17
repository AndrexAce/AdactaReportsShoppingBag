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
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class MainViewModel(
    IProjectFileService projectFileService,
    IDialogService dialogService,
    IStorageService storageService,
    IProductsRepository productsRepository,
    IExcelService excelService,
    INotificationService notificationService)
    : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaveStateText), nameof(SaveButtonVisibility))]
    public partial bool? IsProjectEdited { get; private set; } = null;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsProjectEdited), nameof(NavigationViewMenuItems), nameof(SurveyMenuItemVisibility))]
    public partial ReportPrj? ReportProject { get; private set; } = null;

    public string SaveStateText => IsProjectEdited switch
    {
        true => "– Modifiche non salvate",
        false => "– Nessuna modifica non salvata",
        _ => string.Empty
    };

    public Visibility SaveButtonVisibility =>
        IsProjectEdited is null or false ? Visibility.Collapsed : Visibility.Visible;

    public Visibility SurveyMenuItemVisibility => ReportProject is null ? Visibility.Collapsed : Visibility.Visible;

    public ObservableCollection<Product>? NavigationViewMenuItems => new(ReportProject?.Products ?? []);

    private string? _projectFilePath;

    [RelayCommand]
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

        var project = new ReportPrj(Assembly.GetExecutingAssembly().GetName().Version?.ToString(), projectName, projectCode, products);

        _projectFilePath = projectFileService.CreateProjectFolder(project, userChosenFolder.Path);

        if (_projectFilePath is null) return;

        ReportProject = project;
        IsProjectEdited = false;
    }

    [RelayCommand]
    private async Task OpenProjectAsync()
    {
        var file = await dialogService.ShowFileOpenPickerAsync(".reportprj", "AdactaReportsShoppingBagOpenProjectPicker");

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

    [RelayCommand]
    private async Task OpenClassesFileAsync()
    {
        if (ReportProject is null) return;

        try
        {
            var projectFolderPath = Path.GetDirectoryName(_projectFilePath) ?? throw new FileNotFoundException("The project folder path could not be reached.");
            var excelFilePath = Path.Combine(projectFolderPath, $"Classi{ReportProject.ProjectCode}.xlsx");

            // Open the file with the default associated application
            Process.Start(new ProcessStartInfo
            {
                FileName = excelFilePath,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Maximized
            });
        }
        catch
        {
            await dialogService.ShowInformationDialogAsync("Errore apertura file classi", "Il file non esiste, è danneggiato o non hai i permessi necessari.", "Ok");
        }
    }

    [RelayCommand]
    private async Task ImportInputFileAsync()
    {
        var file = await dialogService.ShowFileOpenPickerAsync(".xlsx", "AdactaReportsShoppingBagOpenInputFilePicker");

        if (file == null) return;

        var notificationId = notificationService.ShowProgressNotification("Importazione file in corso...", "Potrebbero volerci alcuni minuti.");

        await excelService.ImportSurveyFile(file, notificationId);
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

    partial void OnReportProjectChanged(ReportPrj? oldValue, ReportPrj? newValue)
    {
        // Unsubscribe from old products
        if (oldValue?.Products is not null)
        {
            foreach (var product in oldValue.Products.OfType<INotifyPropertyChanged>())
            {
                product.PropertyChanged -= Product_PropertyChanged;

                // Unsubscribe from product photos
                if (product is Product p)
                {
                    foreach (var photo in p.ProductPhotos.OfType<INotifyPropertyChanged>())
                    {
                        photo.PropertyChanged -= ProductPhoto_PropertyChanged;
                    }
                }
            }
        }

        // Subscribe to new products
        if (newValue?.Products is not null)
        {
            foreach (var product in newValue.Products.OfType<INotifyPropertyChanged>())
            {
                product.PropertyChanged += Product_PropertyChanged;

                // Subscribe to product photos
                if (product is Product p)
                {
                    foreach (var photo in p.ProductPhotos.OfType<INotifyPropertyChanged>())
                    {
                        photo.PropertyChanged += ProductPhoto_PropertyChanged;
                    }
                }
            }
        }
    }

    private void Product_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        IsProjectEdited = true;
    }

    private void ProductPhoto_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        IsProjectEdited = true;
    }

    ~MainViewModel()
    {
        // Unsubscribe property changed from products
        if (ReportProject?.Products is not null)
        {
            foreach (var product in ReportProject.Products.OfType<INotifyPropertyChanged>())
            {
                product.PropertyChanged -= Product_PropertyChanged;

                // Unsubscribe from product photos
                if (product is Product p)
                {
                    foreach (var photo in p.ProductPhotos.OfType<INotifyPropertyChanged>())
                    {
                        photo.PropertyChanged -= ProductPhoto_PropertyChanged;
                    }
                }
            }
        }
    }
}