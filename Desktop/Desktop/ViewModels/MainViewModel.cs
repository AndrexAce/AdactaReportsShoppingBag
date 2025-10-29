using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Reflection;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class MainViewModel(IProjectFileService projectFileService, IDialogService dialogService)
    : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaveStateText), nameof(SaveButtonVisibility))]
    public partial bool? IsProjectEdited { get; private set; } = null;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsProjectEdited))]
    public partial ReportPrj? ReportProject { get; private set; } = null;

    public string SaveStateText => IsProjectEdited switch
    {
        true => "– Modifiche non salvate",
        false => "– Nessuna modifica non salvata",
        _ => string.Empty
    };

    public Visibility SaveButtonVisibility =>
        IsProjectEdited is null or false ? Visibility.Collapsed : Visibility.Visible;

    private string? _projectFilePath;

    [RelayCommand]
    private async Task NewProjectAsync()
    {
        var (choice, projectCode, projectName) = await dialogService.ShowDoubleTextboxDialogAsync("Crea nuovo progetto",
                "Crea", "Annulla", "Codice", "Nome (visualizzato su PowerPoint)");

        if (choice is not ContentDialogResult.Primary || projectCode is null || projectName is null) return;

        var userChosenFolder = await dialogService.ShowFolderPicker();

        if (userChosenFolder is null) return;

        var project = new ReportPrj
        {
            ProjectCode = projectCode,
            ProjectName = projectName,
            Version = Assembly.GetExecutingAssembly().GetName().Version?.ToString()
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
    private async Task OpenInfoAsync()
    {
        await dialogService.ShowCreditsDialogAsync();
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