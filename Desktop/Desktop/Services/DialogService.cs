using AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.Generic;
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

    public async Task<(ContentDialogResult, string, string)> ShowDoubleTextboxDialogAsync(string title, string confirmButtonText,
        string cancelButtonText, string firstLabel, string secondLabel)
    {
        var control = new DoubleTextboxControl
        {
            FirstLabel = firstLabel,
            SecondLabel = secondLabel,
        };

        var dialog = new ContentDialog
        {
            Title = title,
            PrimaryButtonText = confirmButtonText,
            SecondaryButtonText = cancelButtonText,
            Content = control,
            XamlRoot = _window?.Content.XamlRoot
        };

        return (await dialog.ShowAsync(), control.FirstValue, control.SecondValue);
    }

    public async Task<StorageFile?> ShowFileOpenPickerAsync()
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

    public async Task<StorageFile?> ShowFileSavePickerAsync(string projectCode)
    {
        FileSavePicker savePicker = new()
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            SuggestedFileName = projectCode,
            CommitButtonText = "Salva",
            DefaultFileExtension = ".reportprj",
            SettingsIdentifier = "AdactaReportsShoppingBagCreateProjectPicker"
        };
        savePicker.FileTypeChoices.Add("Report Project", new List<string> { ".reportprj" });

        var hwnd = WindowNative.GetWindowHandle(_window);
        InitializeWithWindow.Initialize(savePicker, hwnd);

        return await savePicker.PickSaveFileAsync();
    }
}