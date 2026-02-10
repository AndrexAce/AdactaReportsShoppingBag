using System.Reflection;
using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class ProjectFileService(IExcelService excelService) : IProjectFileService
{
    public async Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile)
    {
        return await IsProjectFileValidAsync(projectFile) switch
        {
            (false, _) => null,
            (true, null) => null,
            (true, { } project) => project
        };
    }

    public Task SaveProjectFileAsync(ReportPrj project, string projectFilePath)
    {
        var projectJson = JObject.FromObject(project).ToString();
        return File.WriteAllTextAsync(projectFilePath, projectJson);
    }

    public async Task<string?> CreateProjectFolderAsync(ReportPrj project, string folderPath)
    {
        try
        {
            var projectFolderPath = Path.Combine(folderPath, $"ShoppingBag{project.ProjectCode}");

            Directory.CreateDirectory(projectFolderPath);

            var projectFilePath = Path.Combine(projectFolderPath, $"{project.ProjectCode}.reportprj");

            await excelService.CreateClassesFileAsync(project, projectFolderPath);
            await excelService.CreateSurveyDataFileAsync(project, projectFolderPath);

            await SaveProjectFileAsync(project, projectFilePath);

            return projectFilePath;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<(bool, ReportPrj?)> IsProjectFileValidAsync(IStorageFile file)
    {
        // Validate the file type
        if (file.FileType != ".reportprj") return (false, null);

        try
        {
            // Validate file content
            await using var schemaStream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(
                    "AdactaInternational.AdactaReportsShoppingBag.Desktop.Assets.ReportPrj.schema.json");

            if (schemaStream is null) return (false, null);

            using var schemaStreamReader = new StreamReader(schemaStream);
            var schemaJson = await schemaStreamReader.ReadToEndAsync();
            var projectJson = await File.ReadAllTextAsync(file.Path);

            var schema = JSchema.Parse(schemaJson);
            var project = JObject.Parse(projectJson);

            // If the project is valid, deserialize it to ReportPrj
            return project.IsValid(schema) ? (true, project.ToObject<ReportPrj>()) : (false, null);
        }
        catch
        {
            return (false, null);
        }
    }
}