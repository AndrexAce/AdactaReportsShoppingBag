using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IDialogService
{
    public void SetWindow(Window window);

    public Task<ContentDialogResult> ShowInformationDialogAsync(string title, string content, string closeButtonText);

    public Task<ContentDialogResult> ShowConfirmationDialogAsync(string title, string content, string closeButtonText, string confirmButtonText, string cancelButtonText);

    public Task<StorageFile?> ShowFilePickerAsync();
}
