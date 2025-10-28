using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using Windows.Storage;
using Windows.UI;
using Windows.UI.ViewManagement;
using WinRT.Interop;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop;

internal sealed partial class MainWindow : Window
{
    public MainViewModel ViewModel { get; private set; }
    private readonly IStorageFile? _projectFile;
    private readonly UISettings _uiSettings = new();

    private const string LogoDarkThemePath = "Assets/LogoDarkTheme.png";
    private const string LogoLightThemePath = "Assets/LogoLightTheme.png";

    public MainWindow(IStorageFile? storageFile)
    {
        InitializeComponent();

        ViewModel = Ioc.Default.GetRequiredService<MainViewModel>();
        _projectFile = storageFile;

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
        var foreground = _uiSettings.GetColorValue(UIColorType.Foreground);
        var isDarkMode = IsColorLight(foreground);

        var uri = new Uri($"ms-appx:///{(isDarkMode ? LogoDarkThemePath : LogoLightThemePath)}");
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

            var uri = new Uri($"ms-appx:///{(isDarkMode ? LogoDarkThemePath : LogoLightThemePath)}");
            TitleBarIcon.Source = new BitmapImage(uri);
        });
    }

    private static bool IsColorLight(Color color)
    {
        int brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
        return brightness > 128;
    }

    public async void RootFrame_Loaded(object sender, RoutedEventArgs e)
    {
        if (_projectFile != null)
        {
            await ViewModel.LoadProjectFileAsync(_projectFile);
        }
    }
}