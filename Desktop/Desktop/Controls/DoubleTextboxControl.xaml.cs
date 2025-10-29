using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;

internal sealed partial class DoubleTextboxControl : UserControl
{
    public static readonly DependencyProperty FirstLabelProperty =
        DependencyProperty.Register(
            nameof(FirstLabel),
            typeof(string),
            typeof(DoubleTextboxControl),
            new PropertyMetadata(string.Empty));

    public static readonly DependencyProperty SecondLabelProperty =
        DependencyProperty.Register(
            nameof(SecondLabel),
            typeof(string),
            typeof(DoubleTextboxControl),
            new PropertyMetadata(string.Empty));

    public static readonly DependencyProperty FirstValueProperty =
        DependencyProperty.Register(
            nameof(FirstValue),
            typeof(string),
            typeof(DoubleTextboxControl),
            new PropertyMetadata(string.Empty));

    public static readonly DependencyProperty SecondValueProperty =
        DependencyProperty.Register(
            nameof(SecondValue),
            typeof(string),
            typeof(DoubleTextboxControl),
            new PropertyMetadata(string.Empty));

    public string FirstLabel
    {
        get => (string)GetValue(FirstLabelProperty);
        set => SetValue(FirstLabelProperty, value);
    }

    public string SecondLabel
    {
        get => (string)GetValue(SecondLabelProperty);
        set => SetValue(SecondLabelProperty, value);
    }

    public string FirstValue
    {
        get => (string)GetValue(FirstValueProperty);
        set => SetValue(FirstValueProperty, value);
    }

    public string SecondValue
    {
        get => (string)GetValue(SecondValueProperty);
        set => SetValue(SecondValueProperty, value);
    }

    public DoubleTextboxControl()
    {
        this.InitializeComponent();
    }
}