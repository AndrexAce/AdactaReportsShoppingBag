using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using System.ComponentModel.DataAnnotations;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class PenelopeCredentialsViewModel : ObservableValidator
{
    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [NotifyPropertyChangedFor(nameof(UsernameError), nameof(UsernameErrorVisibility))]
    [NotifyDataErrorInfo]
    public partial string? Username { get; set; }

    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [NotifyPropertyChangedFor(nameof(PasswordError), nameof(PasswordErrorVisibility))]
    [NotifyDataErrorInfo]
    public partial string? Password { get; set; }

    [ObservableProperty] public partial bool IsConfirmButtonEnabled { get; private set; }

    public string? UsernameError => GetErrors(nameof(Username)).FirstOrDefault()?.ErrorMessage;

    public string? PasswordError => GetErrors(nameof(Password)).FirstOrDefault()?.ErrorMessage;

    public Visibility UsernameErrorVisibility =>
        GetErrors(nameof(Username)).Any() ? Visibility.Visible : Visibility.Collapsed;

    public Visibility PasswordErrorVisibility =>
        GetErrors(nameof(Password)).Any() ? Visibility.Visible : Visibility.Collapsed;

    partial void OnUsernameChanged(string? value)
    {
        ValidateProperty(Password, nameof(Password));
        UpdateConfirmButtonState();
    }

    partial void OnPasswordChanged(string? value)
    {
        ValidateProperty(Username, nameof(Username));
        UpdateConfirmButtonState();
    }

    private void UpdateConfirmButtonState()
    {
        IsConfirmButtonEnabled = !HasErrors;
    }
}