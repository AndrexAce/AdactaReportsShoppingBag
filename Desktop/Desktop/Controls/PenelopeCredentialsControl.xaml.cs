using AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.UI.Xaml;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Controls;

[RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
internal sealed partial class PenelopeCredentialsControl
{
    private readonly PenelopeCredentialsViewModel _viewModel;

    public static readonly DependencyProperty UsernameProperty =
        DependencyProperty.Register(
            nameof(Username),
            typeof(string),
            typeof(NewProjectControl),
            new PropertyMetadata(string.Empty));

    public static readonly DependencyProperty PasswordProperty =
        DependencyProperty.Register(
            nameof(Password),
            typeof(string),
            typeof(NewProjectControl),
            new PropertyMetadata(string.Empty));

    public string Username
    {
        get => _viewModel.Username ?? "";
        private set => _viewModel.Username = value;
    }

    public string Password
    {
        get => _viewModel.Password ?? "";
        private set => _viewModel.Password = value;
    }

    public bool IsConfirmButtonEnabled => _viewModel.IsConfirmButtonEnabled;

    public event PropertyChangedEventHandler? IsConfirmButtonEnabledChanged;

    public PenelopeCredentialsControl()
    {
        InitializeComponent();

        _viewModel = Ioc.Default.GetRequiredService<PenelopeCredentialsViewModel>();

        _viewModel.PropertyChanged += ConfirmButtonEnabled_Changed;
    }

    ~PenelopeCredentialsControl()
    {
        _viewModel.PropertyChanged -= ConfirmButtonEnabled_Changed;
    }

    private void ConfirmButtonEnabled_Changed(object? sender, PropertyChangedEventArgs args)
    {
        if (args.PropertyName == nameof(_viewModel.IsConfirmButtonEnabled))
            IsConfirmButtonEnabledChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsConfirmButtonEnabled)));
    }
}