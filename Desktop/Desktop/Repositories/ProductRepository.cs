using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal sealed class ProductRepository(IPenelopeClient penelopeClient) : IProductsRepository
{
    public async Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        return await penelopeClient.GetProductsAsync(jobCode);
    }
}