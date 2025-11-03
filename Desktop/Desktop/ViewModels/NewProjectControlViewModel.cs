using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal partial class NewProjectControlViewModel : ObservableValidator
{
    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [RegularExpression(@"^\d{2}\.\d{3}[a-zA-Z]?$", ErrorMessage = "Formato non valido.")]
    [NotifyPropertyChangedFor(nameof(ProjectCodeError), nameof(ProjectCodeErrorVisibility), nameof(IsConfirmButtonEnabled))]
    [NotifyDataErrorInfo]
    public partial string ProjectCode { get; set; }

    [ObservableProperty]
    [Required(ErrorMessage = "Inserire un valore.")]
    [StringLength(100, MinimumLength = 1, ErrorMessage = "Inserire meno di 100 caratteri.")]
    [NotifyPropertyChangedFor(nameof(ProjectNameError), nameof(ProjectNameErrorVisibility), nameof(IsConfirmButtonEnabled))]
    [NotifyDataErrorInfo]
    public partial string ProjectName { get; set; }

    [ObservableProperty]
    public partial bool IsConfirmButtonEnabled { get; private set; }

    public string? ProjectCodeError => GetErrors(nameof(ProjectCode)).FirstOrDefault()?.ErrorMessage;

    public string? ProjectNameError => GetErrors(nameof(ProjectName)).FirstOrDefault()?.ErrorMessage;

    public Visibility ProjectCodeErrorVisibility => GetErrors(nameof(ProjectCode)).Any() ? Visibility.Visible : Visibility.Collapsed;

    public Visibility ProjectNameErrorVisibility => GetErrors(nameof(ProjectName)).Any() ? Visibility.Visible : Visibility.Collapsed;

    protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        base.OnPropertyChanged(e);

        if (e.PropertyName == nameof(HasErrors))
        {
            IsConfirmButtonEnabled = !HasErrors;
        }
    }
}