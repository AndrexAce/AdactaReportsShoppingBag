using System.Reflection;
using Microsoft.UI.Xaml;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;

internal sealed partial class CreditsControl
{
    public static readonly DependencyProperty VersionProperty =
        DependencyProperty.Register(
            nameof(Version),
            typeof(string),
            typeof(CreditsControl),
            new PropertyMetadata(Assembly.GetExecutingAssembly().GetName().Version?.ToString()));

    public static readonly DependencyProperty YearProperty =
        DependencyProperty.Register(
            nameof(Year),
            typeof(string),
            typeof(CreditsControl),
            new PropertyMetadata(DateTime.Now.Year.ToString()));

    public CreditsControl()
    {
        InitializeComponent();
    }

    public string Version
    {
        get => (string)GetValue(VersionProperty);
        set => SetValue(VersionProperty, value);
    }

    public string Year
    {
        get => (string)GetValue(YearProperty);
        set => SetValue(YearProperty, value);
    }
}