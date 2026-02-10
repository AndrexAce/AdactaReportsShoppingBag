using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IExcelService
{
    public Task CreateClassesFileAsync(ReportPrj project, string projectFolderPath);

    public Task CreateSurveyDataFileAsync(ReportPrj project, string projectFolderPath);

    public Task ImportPenelopeFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, ICollection<Product> products);

    public Task CreateProductFilesAsync(Guid notificationId, ICollection<Product> products, string projectFolderPath,
        string projectCode);

    public Task ProcessProductFilesAsync(Guid notificationId, ICollection<string> fileNames);
}