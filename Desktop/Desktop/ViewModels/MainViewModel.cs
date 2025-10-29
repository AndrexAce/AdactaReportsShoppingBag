using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
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
        // TODO
    }

    [RelayCommand]
    private async Task OpenProjectAsync()
    {
        var file = await dialogService.ShowFilePickerAsync();

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
        // TODO
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