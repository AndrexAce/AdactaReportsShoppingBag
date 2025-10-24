using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI;
using Windows.UI.ViewManagement;
using WinRT.Interop;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop;

internal sealed partial class MainWindow : Window
{
    public MainViewModel MainViewModel { get; } = new();
    private readonly UISettings _uiSettings = new();

    public MainWindow()
    {
        InitializeComponent();

        InitializeAppWindowAndSize(640, 480);
    }

    private void InitializeAppWindowAndSize(int width, int height)
    {
        // Set the window icon
        AppWindow.SetIcon("Assets/favicon.ico");

        // Get the window
        var hwnd = WindowNative.GetWindowHandle(this);
        var windowId = Win32Interop.GetWindowIdFromWindow(hwnd);
        var appWindow = AppWindow.GetFromWindowId(windowId);

        // Resize the app window
        appWindow?.Resize(new Windows.Graphics.SizeInt32 { Width = width, Height = height });

        // Ensure the default presenter is applied, then configure it
        appWindow?.SetPresenter(AppWindowPresenterKind.Default);

        if (appWindow?.Presenter is OverlappedPresenter overlapped)
        {
            overlapped.PreferredMinimumWidth = 640;
            overlapped.PreferredMinimumHeight = 480;
        }

        // Dissolve the system title bar
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TitleBar);

        // Show an icon based on the system theme and keep it updated when the system theme changes
        ShowIconBasedOnSystemTheme();
    }

    private void ShowIconBasedOnSystemTheme()
    {
        // Apply initial icon according to current UI foreground color
        var foreground = _uiSettings.GetColorValue(UIColorType.Foreground);
        var isDarkMode = IsColorLight(foreground);

        var uri = new Uri(isDarkMode
                ? "ms-appx:///Assets/LogoDarkTheme.png"
                : "ms-appx:///Assets/LogoLightTheme.png");
        TitleBarIcon.Source = new BitmapImage(uri);

        // Listen for theme/color changes and update the icon on the UI thread
        _uiSettings.ColorValuesChanged += UiSettings_ColorValuesChanged;
    }

    private void UiSettings_ColorValuesChanged(UISettings sender, object args)
    {
        _ = DispatcherQueue.TryEnqueue(() =>
        {
            var foreground = sender.GetColorValue(UIColorType.Foreground);
            var isDarkMode = IsColorLight(foreground);

            var uri = new Uri(isDarkMode
                ? "ms-appx:///Assets/LogoDarkTheme.png"
                : "ms-appx:///Assets/LogoLightTheme.png");
            TitleBarIcon.Source = new BitmapImage(uri);
        });
    }

    private static bool IsColorLight(Color color)
    {
        int brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
        return brightness > 128;
    }

    private async Task ShowInvalidProjectDialogAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "File progetto non valido",
            Content = "Si è verificato un errore cercando di leggere il file del progetto. Potrebbe essere danneggiato.",
            CloseButtonText = "Ok",
            XamlRoot = Content.XamlRoot
        };

        await dialog.ShowAsync();
    }

    private async void RootFrame_Loaded(object sender, RoutedEventArgs args)
    {
        RootFrame.Loaded -= RootFrame_Loaded;

        if (MainViewModel.IsLoaded == false)
        {
            await ShowInvalidProjectDialogAsync();
        }
    }

    private async void OpenProjectButton_Click(object sender, RoutedEventArgs e)
    {
        FileOpenPicker openPicker = new()
        {
            ViewMode = PickerViewMode.Thumbnail,
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            FileTypeFilter = { ".reportprj" },
            SettingsIdentifier = "AdactaReportsShoppingBagOpenProjectPicker"
        };

        var hwnd = WindowNative.GetWindowHandle(this);
        InitializeWithWindow.Initialize(openPicker, hwnd);

        StorageFile file = await openPicker.PickSingleFileAsync();

        if (file != null)
        {
            MainViewModel.LoadProjectFile(file);

            if (MainViewModel.IsLoaded == false)
            {
                await ShowInvalidProjectDialogAsync();
            }
        }
    }

    private async void NewProjectButton_Click(object sender, RoutedEventArgs e)
    {
        // TODO
    }

    private async void SaveProjectButton_Click(object sender, RoutedEventArgs e)
    {
        // TODO
    }

    private async void HelpButton_Click(object sender, RoutedEventArgs e)
    {
        // TODO
    }
}