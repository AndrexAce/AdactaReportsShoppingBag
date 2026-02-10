using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal sealed class ProductRepository(IPenelopeClient penelopeClient) : IProductsRepository
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        return penelopeClient.GetProductsAsync(jobCode);
    }
}