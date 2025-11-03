using AdactaInternational.AdactaReportsShoppingBag.Model;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IPenelopeClient
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}