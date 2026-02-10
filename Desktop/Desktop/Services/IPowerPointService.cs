using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IPowerPointService
{
    public Task CreateProductSlideshowAsync(Guid notificationId, ICollection<Product> products,
        string projectFolderPath, string projectName, string projectCode);
}