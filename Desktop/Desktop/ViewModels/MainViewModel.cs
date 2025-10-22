using AdactaInternational.AdactaReportsShoppingBag.Desktop.Project;
using CommunityToolkit.Mvvm.ComponentModel;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal partial class MainViewModel : ObservableObject
{
    private ReportPrj? _reportPrj;

    public ReportPrj? ReportPrj
    {
        get => _reportPrj;
        set => SetProperty(ref _reportPrj, value);
    }
}
