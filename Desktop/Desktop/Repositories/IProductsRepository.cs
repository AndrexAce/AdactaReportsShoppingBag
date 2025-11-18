using System.Collections.Generic;
using System.Threading.Tasks;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Repositories;

internal interface IProductsRepository
{
    public Task<IEnumerable<Product>> GetProductsAsync(string jobCode);
}