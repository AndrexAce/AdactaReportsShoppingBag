using AdactaInternational.AdactaReportsShoppingBag.Model;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class PenelopeClient : IPenelopeClient
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        throw new NotImplementedException();
    }
}
