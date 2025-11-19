using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

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
        ExcelApp = new Application
        {
            Visible = false,
            DisplayAlerts = false
        };
        Workbooks = ExcelApp.Workbooks;
        Workbook = Workbooks.Add();
        Sheets = Workbook.Sheets;

        for (var i = 0; i < project.Products.Count(); i++)
        {
            Worksheet worksheet;

            if (i == 0)
                // Use the first default sheet
                worksheet = (Worksheet)Sheets[1];
            else
                // Add new sheet after the last one
                worksheet = (Worksheet)Sheets.Add(After: Worksheets[^1]);

            // Rename the worksheet to match the product code
            worksheet.Name = project.Products.ElementAt(i).Code;

            // Add the worksheet to the collection
            Worksheets.Add(worksheet);
        }

        Workbook.SaveAs(excelFilePath);
    }

    #endregion

    #region Survey file import

    public async Task ImportSurveyFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath)
    {
        await Task.Run(async () => await ExecuteWithCleanupAsync(async () =>
            await ImportSurveyFileInternal(storageFile, notificationId, projectCode, projectFolderPath)));
    }

    private async Task ImportSurveyFileInternal(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Sheets? classesSheets = null;
        Collection<Worksheet> classesWorksheets = [];
        Range? responseTableRange = null;

        try
        {
            // Create a silent Excel application
            ExcelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            Workbooks = ExcelApp.Workbooks;
            Workbook = Workbooks.Open(storageFile.Path);
            Sheets = Workbook.Sheets;

            // Open the survey classes file
            classesWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;

            // For each worksheet in the survey file, find the corresponding worksheet in the classes file and populate it
            foreach (Worksheet sheet in Sheets)
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
                Worksheets.Add(sheet);
                classesWorksheets.Add(classesWorksheet);
            }

            classesWorkbook.Save();

            await notificationService.RemoveNotification(notificationId);

            notificationService.ShowProgressNotification("Importazione completata",
                "Il file è stato importato con successo.");
        }
        finally // Clean up the resources not managed by the base class
        {
            if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);

            foreach (var element in classesWorksheets) Marshal.ReleaseComObject(element);
            Worksheets.Clear();

            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (classesWorkbook is not null)
            {
                classesWorkbook.Close(false);
                Marshal.ReleaseComObject(classesWorkbook);
            }
        }
    }

    private static DataTable AddQuestionsColumn(DataTable oldDataTable)
    {
        // Create the new datatable with the questions column
        var newDataTable = new DataTable();
        newDataTable.Columns.Add(new DataColumn("Domanda", typeof(string)));

        // Extract the questions from the original datatable's column names (the Excel header row)
        var questions = oldDataTable.Columns
            .Cast<DataColumn>()
            .Select(c => c.ColumnName)
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

    private static DataTable AddFieldNameColumn(DataTable newDataTable)
    {
        // Create the field names column
        newDataTable.Columns.Add(new DataColumn("Etichetta", typeof(string)));

        return newDataTable;
    }

    private static DataTable AddCategoryColumn(DataTable newDataTable)
    {
        // Create the category column
        newDataTable.Columns.Add(new DataColumn("Classe", typeof(string)));

        return newDataTable;
    }

    #endregion

    #region Classes file import

    public async Task ImportClassesFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, string productCode)
    {
        await Task.Run(async () => await ExecuteWithCleanupAsync(async () =>
            await ImportClassesFileInternal(storageFile, notificationId, projectCode, projectFolderPath, productCode)));
    }

    private async Task ImportClassesFileInternal(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, string productCode)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Sheets? classesSheets = null;
        Worksheet? classesSheet = null;
        Range? responseTableRange = null;

        try
        {
            // Create a silent Excel application
            ExcelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            Workbooks = ExcelApp.Workbooks;
            Workbook = Workbooks.Open(storageFile.Path);
            Sheets = Workbook.Sheets;

            // Open the app's survey classes file
            classesWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;

            // Find ActiveViewing's classes file and import it
            foreach (Worksheet sheet in Sheets)
            {
                if (!sheet.Name.Contains("Classi domande"))
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                // Find the corresponding worksheet in the app's classes file.
                // If there is none, create the sheet.
                try
                {
                    classesSheet = classesSheets.Item[productCode];
                }
                catch
                {
                    classesSheet = classesSheets.Add();
                    classesSheet.Name = productCode;
                }

                // Copy and paste the table
                responseTableRange = sheet.UsedRange;

                var originalDataTable = responseTableRange.MakeDataTable();

                // Step 1: Remove the useless rows
                var newDataTable = originalDataTable.RemoveLastRows(3);

                // Step 2: Remove the useless columns
                newDataTable = newDataTable.RemoveLastColumns(6);

                // Step 3 : Take the needed columns and rename them
                newDataTable = TakeColumnsAndRename(newDataTable);

                // Step 4: Write the new datatable to the classes worksheet
                newDataTable.WriteToWorksheet(classesSheet, "Classi");

                Marshal.ReleaseComObject(sheet);
                break;
            }

            classesWorkbook.Save();

            await notificationService.RemoveNotification(notificationId);

            notificationService.ShowProgressNotification("Importazione completata",
                "Il file è stato importato con successo.");
        }
        finally // Clean up the resources not managed by the base class
        {
            if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);

            if (classesSheet is not null) Marshal.ReleaseComObject(classesSheet);

            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (classesWorkbook is not null)
            {
                classesWorkbook.Close(false);
                Marshal.ReleaseComObject(classesWorkbook);
            }
        }
    }

    private static DataTable TakeColumnsAndRename(DataTable oldDataTable)
    {
        // Take only the needed columns and rename them
        var newColumns = oldDataTable.Copy().Columns
            .Cast<DataColumn>()
            .Where((c, index) => index is 1 or 4 or 5)
            .Select(c =>
            {
                c.ColumnName = c.ColumnName switch
                {
                    "Testo Domanda" => "Domanda",
                    "Etichetta Domanda" => "Etichetta",
                    _ => c.ColumnName
                };

                return c;
            });

        // Add the data to a new datatable with the new columns
        var newDataTable = new DataTable();

        foreach (var column in newColumns)
            newDataTable.Columns.Add(new DataColumn(column.ColumnName, typeof(string)));

        foreach (DataRow row in oldDataTable.Rows)
        {
            var newRow = newDataTable.NewRow();
            newRow["Domanda"] = row["Testo Domanda"];
            newRow["Etichetta"] = row["Etichetta Domanda"];
            newRow["Classe"] = row["Classe"];
            newDataTable.Rows.Add(newRow);
        }

        return newDataTable;
    }

    #endregion
}