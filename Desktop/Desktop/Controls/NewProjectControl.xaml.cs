using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.ComponentModel;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;

internal sealed partial class NewProjectControl : UserControl
{
    private readonly NewProjectControlViewModel _viewModel;

    public static readonly DependencyProperty ProjectCodeProperty =
        DependencyProperty.Register(
            nameof(ProjectCode),
            typeof(string),
            typeof(NewProjectControl),
            new PropertyMetadata(string.Empty));

    public static readonly DependencyProperty ProjectNameProperty =
        DependencyProperty.Register(
            nameof(ProjectName),
            typeof(string),
            typeof(NewProjectControl),
            new PropertyMetadata(string.Empty));

    public string ProjectCode
    {
        get => _viewModel.ProjectCode;
        private set => _viewModel.ProjectCode = value;
    }

    public string ProjectName
    {
        get => _viewModel.ProjectName;
        private set => _viewModel.ProjectName = value;
    }

    public bool IsConfirmButtonEnabled => _viewModel.IsConfirmButtonEnabled;

    public event PropertyChangedEventHandler? IsConfirmButtonEnabledChanged;

    public NewProjectControl()
    {
        InitializeComponent();

        _viewModel = Ioc.Default.GetRequiredService<NewProjectControlViewModel>();

        _viewModel.PropertyChanged += ConfirmButtonEnabled_Changed;
    }

    ~NewProjectControl()
    {
        _viewModel.PropertyChanged -= ConfirmButtonEnabled_Changed;
    }

    private void ConfirmButtonEnabled_Changed(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(_viewModel.IsConfirmButtonEnabled))
        {
            IsConfirmButtonEnabledChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsConfirmButtonEnabled)));
        }
    }
}