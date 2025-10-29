using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IDialogService
{
    public void SetWindow(Window window);

    public Task<ContentDialogResult> ShowInformationDialogAsync(string title, string content, string closeButtonText);

    public Task<(ContentDialogResult, string, string)> ShowDoubleTextboxDialogAsync(string title, string confirmButtonText, string cancelButtonText, string firstLabel, string secondLabel);

    public Task<StorageFile?> ShowFileOpenPickerAsync();

    public Task<StorageFolder?> ShowFolderPicker();
}