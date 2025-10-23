using AdactaInternational.AdactaReportsShoppingBag.Desktop.Project;
using CommunityToolkit.Mvvm.ComponentModel;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal partial class MainViewModel : ObservableObject
{
    public bool? IsLoaded { get; set; }

    [ObservableProperty]
    public partial ReportPrj? ReportPrj { get; set; }
}
