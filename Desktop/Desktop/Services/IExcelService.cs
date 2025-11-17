using AdactaInternational.AdactaReportsShoppingBag.Model;
using System;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IExcelService
{
    public void CreateClassesFile(ReportPrj project, string projectFolderPath);

    public Task ImportSurveyFile(IStorageFile storageFile, Guid notificationId);
}
