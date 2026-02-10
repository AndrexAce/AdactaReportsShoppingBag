using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IPenelopeClient
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}