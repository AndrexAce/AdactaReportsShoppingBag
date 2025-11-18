using System.Threading.Tasks;
using Windows.Storage;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IDialogService
{
    public void SetWindow(Window window);

    public Task<ContentDialogResult> ShowInformationDialogAsync(string title, string content, string closeButtonText);

    public Task<(ContentDialogResult, string, string)> ShowNewProjectDialogAsync(string title, string confirmButtonText,
        string cancelButtonText);

    public Task<(ContentDialogResult, string, string)> ShowPenelopeCredentialsDialogAsync(string title,
        string confirmButtonText,
        string cancelButtonText);

    public Task<StorageFile?> ShowFileOpenPickerAsync(string fileExtension, string settingsIdentifier);

    public Task<StorageFolder?> ShowFolderPicker();

    public Task ShowCreditsDialogAsync();

    public Task ShowImageDialogAsync(string imageUrl);
}