using AdactaInternational.AdactaReportsShoppingBag.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class ExcelService(INotificationService notificationService) : ExcelComHandler, IExcelService
{
    public void CreateClassesFile(ReportPrj project, string projectFolderPath)
    {
        ExecuteWithCleanup(() => CreateClassesFileInternal(project, projectFolderPath));
    }

    private void CreateClassesFileInternal(ReportPrj project, string projectFolderPath)
    {
        var excelFilePath = Path.Combine(projectFolderPath, $"Classi{project.ProjectCode}.xlsx");

        // Create a silent Excel application
        excelApp = new Application
        {
            Visible = false,
            DisplayAlerts = false
        };
        workbooks = excelApp.Workbooks;
        workbook = workbooks.Add();
        sheets = workbook.Sheets;

        for (int i = 0; i < project.Products.Count(); i++)
        {
            Worksheet worksheet;

            if (i == 0)
            {
                // Use the first default sheet
                worksheet = (Worksheet)sheets[1];
            }
            else
            {
                // Add new sheet after the last one
                worksheet = (Worksheet)sheets.Add(After: worksheets[^1]);
            }

            // Rename the worksheet to match the product code
            worksheet.Name = project.Products.ElementAt(i).Code;

            // Add the worksheet to the collection
            worksheets.Add(worksheet);
        }

        workbook.SaveAs(excelFilePath);
    }

    public async Task ImportSurveyFile(IStorageFile storageFile, Guid notificationId)
    {
        await Task.Run(async () => await ExecuteWithCleanupAsync(async () => await ImportSurveyFileInternal(storageFile, notificationId)));
    }

    private async Task ImportSurveyFileInternal(IStorageFile storageFile, Guid notificationId)
    {
        // Create a silent Excel application
        excelApp = new Application
        {
            Visible = false,
            DisplayAlerts = false
        };
        workbooks = excelApp.Workbooks;
        workbook = workbooks.Open(storageFile.Path);
        sheets = workbook.Sheets;

        foreach (Worksheet sheet in sheets)
        {
            worksheets.Add(sheet);

            // TODO: Implement the actual import logic here
        }

        await notificationService.RemoveNotification(notificationId);

        notificationService.ShowProgressNotification("Importazione completata", "Il file è stato importato con successo.");
    }
}
