using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Pages;

internal sealed partial class ProductPage : Page
{
    public Product? CurrentProduct { get; private set; }
    private readonly IDialogService _dialogService;

    public ProductPage()
    {
        InitializeComponent();

        _dialogService = Ioc.Default.GetRequiredService<IDialogService>();
    }

    protected override void OnNavigatedTo(NavigationEventArgs e)
    {
        base.OnNavigatedTo(e);

        if (e.Parameter is Product product)
        {
            CurrentProduct = product;
        }
    }

    private async void Button_Click(object sender, Microsoft.UI.Xaml.RoutedEventArgs args)
    {
        if (sender is Button button && button.Tag is string photoUrl)
        {
            await _dialogService.ShowImageDialogAsync(photoUrl);
        }
    }
}
