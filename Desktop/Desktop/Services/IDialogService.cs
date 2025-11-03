using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IDialogService
{
    public void SetWindow(Window window);

    public Task<ContentDialogResult> ShowInformationDialogAsync(string title, string content, string closeButtonText);

    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    public Task<(ContentDialogResult, string, string)> ShowNewProjectDialogAsync(string title, string confirmButtonText,
        string cancelButtonText);

    public Task<StorageFile?> ShowFileOpenPickerAsync();

    public Task<StorageFolder?> ShowFolderPicker();

    public Task ShowCreditsDialogAsync();
}