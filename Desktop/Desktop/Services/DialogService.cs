using AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Diagnostics.CodeAnalysis;
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

    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    public async Task<(ContentDialogResult, string, string)> ShowNewProjectDialogAsync(string title,
        string confirmButtonText,
        string cancelButtonText)
    {
        var control = new NewProjectControl();

        var dialog = new ContentDialog
        {
            Title = title,
            PrimaryButtonText = confirmButtonText,
            IsPrimaryButtonEnabled = control.IsConfirmButtonEnabled,
            CloseButtonText = cancelButtonText,
            Content = control,
            XamlRoot = _window?.Content.XamlRoot
        };

        control.IsConfirmButtonEnabledChanged += (_, _) =>
        {
            dialog.IsPrimaryButtonEnabled = control.IsConfirmButtonEnabled;
        };

        return (await dialog.ShowAsync(), control.ProjectCode, control.ProjectName);
    }

    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    public async Task<(ContentDialogResult, string, string)> ShowPenelopeCredentialsDialogAsync(string title,
    string confirmButtonText,
    string cancelButtonText)
    {
        var control = new PenelopeCredentialsControl();

        var dialog = new ContentDialog
        {
            Title = title,
            PrimaryButtonText = confirmButtonText,
            IsPrimaryButtonEnabled = control.IsConfirmButtonEnabled,
            CloseButtonText = cancelButtonText,
            Content = control,
            XamlRoot = _window?.Content.XamlRoot
        };

        control.IsConfirmButtonEnabledChanged += (_, _) =>
        {
            dialog.IsPrimaryButtonEnabled = control.IsConfirmButtonEnabled;
        };

        return (await dialog.ShowAsync(), control.Username, control.Password);
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

    public async Task<StorageFolder?> ShowFolderPicker()
    {
        FolderPicker folderPicker = new()
        {
            ViewMode = PickerViewMode.Thumbnail,
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            FileTypeFilter = { ".reportprj" },
            SettingsIdentifier = "AdactaReportsShoppingBagCreateProjectPicker"
        };

        var hwnd = WindowNative.GetWindowHandle(_window);
        InitializeWithWindow.Initialize(folderPicker, hwnd);

        return await folderPicker.PickSingleFolderAsync();
    }

    public async Task ShowCreditsDialogAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "Adacta Reports Shopping Bag",
            CloseButtonText = "Chiudi",
            Content = new CreditsControl(),
            XamlRoot = _window?.Content.XamlRoot
        };

        await dialog.ShowAsync();
    }
}