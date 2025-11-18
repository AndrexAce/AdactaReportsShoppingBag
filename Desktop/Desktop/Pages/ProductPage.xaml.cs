using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Pages;

internal sealed partial class ProductPage
{
    private readonly IDialogService _dialogService;

    public ProductPage()
    {
        InitializeComponent();

        _dialogService = Ioc.Default.GetRequiredService<IDialogService>();
    }

    public Product? CurrentProduct { get; private set; }

    protected override void OnNavigatedTo(NavigationEventArgs e)
    {
        base.OnNavigatedTo(e);

        if (e.Parameter is Product product) CurrentProduct = product;
    }

    private async void Button_Click(object sender, RoutedEventArgs args)
    {
        try
        {
            if (sender is Button { Tag: string photoUrl }) await _dialogService.ShowImageDialogAsync(photoUrl);
        }
        catch
        {
            // Do nothing - errors are already handled
        }
    }
}