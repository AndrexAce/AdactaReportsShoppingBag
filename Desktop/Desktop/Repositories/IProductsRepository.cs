using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal interface IProductsRepository
{
    [RequiresUnreferencedCode("Uses functionality that may break when trimming.")]
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}