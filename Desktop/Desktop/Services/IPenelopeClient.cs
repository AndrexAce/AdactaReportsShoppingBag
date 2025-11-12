using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IPenelopeClient
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}