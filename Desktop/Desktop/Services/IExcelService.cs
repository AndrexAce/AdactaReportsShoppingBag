using System;
using System.Threading.Tasks;
using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Model;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IExcelService
{
    public Task CreateClassesFileAsync(ReportPrj project, string projectFolderPath);

    public Task CreateSurveyDataFileAsync(ReportPrj project, string projectFolderPath);

    public Task ImportPenelopeFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath);

    public Task ImportActiveViewingFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, string productCode);
}