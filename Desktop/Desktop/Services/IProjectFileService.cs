using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Model;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IProjectFileService
{
    public Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile);

    public Task SaveProjectFileAsync(ReportPrj project, string projectFilePath);

    public Task<string?> CreateProjectFolderAsync(ReportPrj project, string folderPath);
}