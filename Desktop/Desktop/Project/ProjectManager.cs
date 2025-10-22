using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using System.IO;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Project;

internal static class ProjectManager
{
    public static bool IsProjectFileValid(IStorageFile file, out ReportPrj? reportPrj)
    {
        reportPrj = null;

        // Validate file type
        if (file.ContentType != "application/json" || file.FileType != ".reportprj")
        {
            return false;
        }

        try
        {
            // Validate file content
            var schemaJson = File.ReadAllText("Assets/Schemas/ProjectSchema.json");
            var projectJson = File.ReadAllText(file.Path);

            JSchema schema = JSchema.Parse(schemaJson);

            JObject project = JObject.Parse(projectJson);

            // If the project is valid, deserialize it to ReportPrj
            if (project.IsValid(schema))
            {
                reportPrj = project.ToObject<ReportPrj>();

                return true;
            }

            return false;
        }
        catch
        {
            return false;
        }
    }
}
