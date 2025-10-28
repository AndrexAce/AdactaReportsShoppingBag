using AdactaInternational.AdactaReportsShoppingBag.Model.Project;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IProjectFileService
{
    public Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile);

    public Task SaveProjectFileAsync(ReportPrj project, string projectFilePath);
}
