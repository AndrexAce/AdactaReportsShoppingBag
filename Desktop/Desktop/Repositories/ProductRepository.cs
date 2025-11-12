using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal sealed class ProductRepository(IPenelopeClient penelopeClient) : IProductsRepository
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        return penelopeClient.GetProductsAsync(jobCode);
    }
}