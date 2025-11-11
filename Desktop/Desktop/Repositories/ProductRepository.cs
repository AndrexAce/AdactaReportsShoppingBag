using AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal sealed class ProductRepository(IPenelopeClient penelopeClient) : IProductsRepository
{
    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        return penelopeClient.GetProductsAsync(jobCode);
    }
}