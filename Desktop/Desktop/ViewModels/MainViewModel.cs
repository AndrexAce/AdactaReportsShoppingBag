using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Project;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class MainViewModel(IProjectFileService _projectFileService, IDialogService _dialogService) : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaveStateText), nameof(SaveButtonVisibility))]
    private partial bool? IsProjectEdited { get; set; } = null;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsProjectEdited))]
    public partial ReportPrj? ReportProject { get; private set; } = null;

    public string SaveStateText => IsProjectEdited switch
    {
        true => "- Modifiche non salvate",
        false => "- Nessuna modifica non salvata",
        _ => string.Empty
    };

    public Visibility SaveButtonVisibility => IsProjectEdited is null or false ? Visibility.Collapsed : Visibility.Visible;

    private string? _projectFilePath = null;

    [RelayCommand]
    private async Task NewProjectAsync()
    {
        // TODO
    }

    [RelayCommand]
    private async Task OpenProjectAsync()
    {
        StorageFile? file = await _dialogService.ShowFilePickerAsync();

        if (file == null) return;

        ReportProject = await _projectFileService.LoadProjectFileAsync(file);

        if (ReportProject is null)
        {
            await _dialogService.ShowInformationDialogAsync("Progetto non caricato", "Il file del progetto è danneggiato.", "Ok");
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

        await _projectFileService.SaveProjectFileAsync(ReportProject, _projectFilePath);

        IsProjectEdited = false;
    }

    public async Task LoadProjectFileAsync(IStorageFile file)
    {
        ReportProject = await _projectFileService.LoadProjectFileAsync(file);

        if (ReportProject is null)
        {
            await _dialogService.ShowInformationDialogAsync("Progetto non caricato", "Il file del progetto è danneggiato.", "Ok");
        }
        else
        {
            _projectFilePath = file.Path;
            IsProjectEdited = false;
        }
    }
}
