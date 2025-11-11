using AdactaInternational.AdactaReportsShoppingBag.Model;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IProjectFileService
{
    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    [RequiresDynamicCode("Uses functionality that may break with AOT.")]
    public Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile);

    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    [RequiresDynamicCode("Uses functionality that may break with AOT.")]
    public Task SaveProjectFileAsync(ReportPrj project, string projectFilePath);

    [RequiresUnreferencedCode("Uses functionality that may break with trimming.")]
    [RequiresDynamicCode("Uses functionality that may break with AOT.")]
    public string? CreateProjectFolder(ReportPrj project, string folderPath);
}