using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class DialogService : IDialogService
{
    private Window? _window;

    public void SetWindow(Window window)
    {
        _window = window;
    }

    public async Task<ContentDialogResult> ShowInformationDialogAsync(string title, string content,
        string closeButtonText)
    {
        var dialog = new ContentDialog
        {
            Title = title,
            Content = content,
            CloseButtonText = closeButtonText,
            XamlRoot = _window?.Content.XamlRoot
        };

        return await dialog.ShowAsync();
    }

    public async Task<StorageFile?> ShowFilePickerAsync()
    {
        FileOpenPicker openPicker = new()
        {
            ViewMode = PickerViewMode.Thumbnail,
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            FileTypeFilter = { ".reportprj" },
            SettingsIdentifier = "AdactaReportsShoppingBagOpenProjectPicker"
        };

        var hwnd = WindowNative.GetWindowHandle(_window);
        InitializeWithWindow.Initialize(openPicker, hwnd);

        return await openPicker.PickSingleFileAsync();
    }
}