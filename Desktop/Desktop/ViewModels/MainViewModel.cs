using AdactaInternational.AdactaReportsShoppingBag.Desktop.Project;
using CommunityToolkit.Mvvm.ComponentModel;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using System.IO;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.ViewModels;

internal partial class MainViewModel : ObservableObject
{
    public bool? IsLoaded { get; private set; } = null;

    [ObservableProperty]
    public partial ReportPrj? ReportPrj { get; private set; } = null;

    public void LoadProjectFile(IStorageFile file)
    {
        IsLoaded = IsProjectFileValid(file, out ReportPrj? reportPrj);

        ReportPrj = reportPrj;
    }

    private static bool IsProjectFileValid(IStorageFile file, out ReportPrj? reportPrj)
    {
        reportPrj = null;

        // Validate file type
        if (file is null || file.FileType != ".reportprj")
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
