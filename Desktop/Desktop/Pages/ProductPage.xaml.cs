using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Pages;

internal sealed partial class ProductPage : Page
{
    public Product? CurrentProduct { get; private set; }

    public ProductPage()
    {
        InitializeComponent();
    }

    protected override void OnNavigatedTo(NavigationEventArgs e)
    {
        base.OnNavigatedTo(e);

        if (e.Parameter is Product product)
        {
            CurrentProduct = product;
        }
    }
}
