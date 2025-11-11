using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IPenelopeClient
{
    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}