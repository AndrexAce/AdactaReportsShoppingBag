using AdactaInternational.AdactaReportsShoppingBag.Model.Project;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using System;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal class ProjectFileService : IProjectFileService
{
    public async Task<ReportPrj?> LoadProjectFileAsync(IStorageFile projectFile)
    {
        return await IsProjectFileValidAsync(projectFile) switch
        {
            (false, _) => null,
            (true, null) => null,
            (true, ReportPrj project) => project
        };
    }

    private static async Task<(bool, ReportPrj?)> IsProjectFileValidAsync(IStorageFile file)
    {
        // Validate file type
        if (file is null || file.FileType != ".reportprj")
        {
            return (false, null);
        }

        try
        {
            // TODO: Load schema from app package resources

            // Validate file content
            var schemaJson = await FileIO.ReadTextAsync(schemaFile).AsTask();
            var projectJson = await FileIO.ReadTextAsync(file).AsTask();

            JSchema schema = JSchema.Parse(schemaJson);

            JObject project = JObject.Parse(projectJson);

            // If the project is valid, deserialize it to ReportPrj
            if (project.IsValid(schema))
            {
                return (true, project.ToObject<ReportPrj>());
            }

            return (false, null);
        }
        catch
        {
            return (false, null);
        }
    }

    public async Task SaveProjectFileAsync(ReportPrj project, string projectFilePath)
    {
        if (project is null || projectFilePath is null)
        {
            return;
        }

        var projectJson = JObject.FromObject(project).ToString();
        await File.WriteAllTextAsync(projectFilePath, projectJson);
    }
}
