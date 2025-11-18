using System;
using System.Threading.Tasks;
using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Model;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IExcelService
{
    public void CreateClassesFile(ReportPrj project, string projectFolderPath);

    public Task ImportSurveyFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath);

    public Task ImportClassesFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath);
}