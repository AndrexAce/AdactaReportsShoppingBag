using AdactaInternational.AdactaReportsShoppingBag.Model;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IProjectFileService
{
    public Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile);

    public Task SaveProjectFileAsync(ReportPrj project, string projectFilePath);

    public Task<string?> CreateProjectFolderAsync(ReportPrj project, string folderPath);
}