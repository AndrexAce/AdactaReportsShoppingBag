using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using System.ComponentModel.DataAnnotations;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal sealed partial class NewProjectControlViewModel : ObservableValidator
{
    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [RegularExpression(@"^[0-9]{2}\.[0-9]{3}[A-Z]?$", ErrorMessage = "Formato non valido.")]
    [NotifyPropertyChangedFor(nameof(ProjectCodeError), nameof(ProjectCodeErrorVisibility))]
    [NotifyDataErrorInfo]
    public partial string? ProjectCode { get; set; }

    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [StringLength(100, MinimumLength = 1, ErrorMessage = "Inserire meno di 100 caratteri.")]
    [NotifyPropertyChangedFor(nameof(ProjectNameError), nameof(ProjectNameErrorVisibility))]
    [NotifyDataErrorInfo]
    public partial string? ProjectName { get; set; }

    [ObservableProperty] public partial bool IsConfirmButtonEnabled { get; private set; }

    public string? ProjectCodeError => GetErrors(nameof(ProjectCode)).FirstOrDefault()?.ErrorMessage;

    public string? ProjectNameError => GetErrors(nameof(ProjectName)).FirstOrDefault()?.ErrorMessage;

    public Visibility ProjectCodeErrorVisibility =>
        GetErrors(nameof(ProjectCode)).Any() ? Visibility.Visible : Visibility.Collapsed;

    public Visibility ProjectNameErrorVisibility =>
        GetErrors(nameof(ProjectName)).Any() ? Visibility.Visible : Visibility.Collapsed;

    partial void OnProjectCodeChanged(string? value)
    {
        ValidateProperty(ProjectName, nameof(ProjectName));
        UpdateConfirmButtonState();
    }

    partial void OnProjectNameChanged(string? value)
    {
        ValidateProperty(ProjectCode, nameof(ProjectCode));
        UpdateConfirmButtonState();
    }

    private void UpdateConfirmButtonState()
    {
        IsConfirmButtonEnabled = !HasErrors;
    }
}