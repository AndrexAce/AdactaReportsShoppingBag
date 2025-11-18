using AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class ExcelService(INotificationService notificationService) : ExcelComHandler, IExcelService
{
    #region Classes file creation
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
    #endregion

    #region Survey file import
    public async Task ImportSurveyFileAsync(IStorageFile storageFile, System.Guid notificationId, string projectCode, string projectFolderPath)
    {
        await Task.Run(async () => await ExecuteWithCleanupAsync(async () => await ImportSurveyFileInternal(storageFile, notificationId, projectCode, projectFolderPath)));
    }

    private async Task ImportSurveyFileInternal(IStorageFile storageFile, System.Guid notificationId, string projectCode, string projectFolderPath)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Sheets? classesSheets = null;
        Collection<Worksheet> classesWorksheets = [];
        Range? responseTableRange = null;

        try
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

            // Open the survey classes file
            classesWorkbook = workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;

            // For each worksheet in the survey file, find the corresponding worksheet in the classes file and populate it
            foreach (Worksheet sheet in sheets)
            {
                Worksheet classesWorksheet = classesSheets.Item[sheet.Name];

                if (classesWorksheet is null) continue;

                // Get the used range of the survey file, it contains the class names
                responseTableRange = sheet.UsedRange;

                var originalDataTable = responseTableRange.MakeDataTable();

                // Step 1: Make the questions column
                var newDataTable = AddQuestionsColumn(originalDataTable);

                // Step 2: Add the field name column
                newDataTable = AddFieldNameColumn(newDataTable);

                // Step 3: Add the category column
                newDataTable = AddCategoryColumn(newDataTable);

                // Step 4: Write the new datatable to the classes worksheet
                newDataTable.WriteToWorksheet(classesWorksheet, "Classi");

                // Add the worksheets to the collection
                worksheets.Add(sheet);
                classesWorksheets.Add(classesWorksheet);
            }

            classesWorkbook.Save();

            await notificationService.RemoveNotification(notificationId);

            notificationService.ShowProgressNotification("Importazione completata", "Il file è stato importato con successo.");
        }
        finally // Clean up the resources not managed by the base class
        {
            if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);

            foreach (var element in classesWorksheets) Marshal.ReleaseComObject(element);
            worksheets.Clear();

            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (classesWorkbook is not null)
            {
                classesWorkbook.Close(false);
                Marshal.ReleaseComObject(classesWorkbook);
            }
        }
    }

    private static System.Data.DataTable AddQuestionsColumn(System.Data.DataTable oldDataTable)
    {
        // Create the new datatable with the questions column
        var newDataTable = new System.Data.DataTable();
        newDataTable.Columns.Add(new System.Data.DataColumn("Domanda", typeof(string)));

        // Extract the questions from the original datatable's column names (the Excel header row)
        var questions = oldDataTable.Columns
            .Cast<System.Data.DataColumn>()
            .Select(c => c.ColumnName ?? string.Empty)
            .Where(s => s.StartsWith('D'))
            .ToArray();

        // Populate the questions datatable
        foreach (var question in questions)
        {
            var dr = newDataTable.NewRow();
            dr[0] = question;
            newDataTable.Rows.Add(dr);
        }

        return newDataTable;
    }

    private static System.Data.DataTable AddFieldNameColumn(System.Data.DataTable newDataTable)
    {
        // Create the field names column
        newDataTable.Columns.Add(new System.Data.DataColumn("Etichetta", typeof(string)));

        return newDataTable;
    }

    private static System.Data.DataTable AddCategoryColumn(System.Data.DataTable newDataTable)
    {
        // Create the category column
        newDataTable.Columns.Add(new System.Data.DataColumn("Classe", typeof(string)));

        return newDataTable;
    }
    #endregion

    #region Classes file import
    public async Task ImportClassesFileAsync(IStorageFile storageFile, System.Guid notificationId, string projectCode, string projectFolderPath)
    {
        await Task.Run(async () => await ExecuteWithCleanupAsync(async () => await ImportClassesFileInternal(storageFile, notificationId, projectCode, projectFolderPath)));
    }

    private async Task ImportClassesFileInternal(IStorageFile storageFile, System.Guid notificationId, string projectCode, string projectFolderPath)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Sheets? classesSheets = null;
        Collection<Worksheet> classesWorksheets = [];
        Range? responseTableRange = null;

        try
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

            // Open the app's survey classes file
            classesWorkbook = workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;

            // Find ActiveViewing's classes file and import it
            foreach (Worksheet sheet in sheets)
            {
                if (!sheet.Name.Contains("Classi domande"))
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                // TODO

                Marshal.ReleaseComObject(sheet);
            }

            classesWorkbook.Save();

            await notificationService.RemoveNotification(notificationId);

            notificationService.ShowProgressNotification("Importazione completata", "Il file è stato importato con successo.");
        }
        finally // Clean up the resources not managed by the base class
        {
            if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);

            foreach (var element in classesWorksheets) Marshal.ReleaseComObject(element);
            worksheets.Clear();

            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (classesWorkbook is not null)
            {
                classesWorkbook.Close(false);
                Marshal.ReleaseComObject(classesWorkbook);
            }
        }
    }
    #endregion
}
