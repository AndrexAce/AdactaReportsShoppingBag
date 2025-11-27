using System;
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

    public async Task CreateClassesFileAsync(ReportPrj project, string projectFolderPath)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                CreateClassesFileInternal(project, projectFolderPath)));
    }

    private void CreateClassesFileInternal(ReportPrj project, string projectFolderPath)
    {
        Worksheet? worksheet = null;

        try
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
            Worksheets = Workbook.Worksheets;

            for (var i = 0; i < project.Products.Count(); i++)
            {
                if (i == 0)
                    // Use the first default sheet
                    worksheet = (Worksheet)Worksheets[1];
                else
                    // Add new sheet after the last one
                    worksheet = (Worksheet)Worksheets.Add();

                // Rename the worksheet to match the product code
                worksheet.Name = project.Products.ElementAt(i).Code;

                // Release the worksheet on each iteration
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }

            Workbook.SaveAs(excelFilePath);
        }
        finally
        {
            if (worksheet is not null) Marshal.ReleaseComObject(worksheet);
        }
    }

    #endregion

    #region Survey data file creation

    public async Task CreateSurveyDataFileAsync(ReportPrj project, string projectFolderPath)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                CreateSurveyDataFileInternal(project, projectFolderPath)));
    }

    private void CreateSurveyDataFileInternal(ReportPrj project, string projectFolderPath)
    {
        Worksheet? worksheet = null;

        try
        {
            var excelFilePath = Path.Combine(projectFolderPath, $"Dati{project.ProjectCode}.xlsx");

            // Create a silent Excel application
            ExcelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            Workbooks = ExcelApp.Workbooks;
            Workbook = Workbooks.Add();
            Worksheets = Workbook.Worksheets;

            for (var i = 0; i < project.Products.Count(); i++)
            {
                if (i == 0)
                    // Use the first default sheet
                    worksheet = (Worksheet)Worksheets[1];
                else
                    // Add new sheet after the last one
                    worksheet = (Worksheet)Worksheets.Add();

                // Rename the worksheet to match the product code
                worksheet.Name = project.Products.ElementAt(i).Code;

                // Release the worksheet on each iteration
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }

            Workbook.SaveAs(excelFilePath);
        }
        finally
        {
            if (worksheet is not null) Marshal.ReleaseComObject(worksheet);
        }
    }

    #endregion

    #region Penelope file import

    public async Task ImportPenelopeFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ImportPenelopeFileInternal(storageFile, notificationId, projectCode, projectFolderPath)));
    }

    private void ImportPenelopeFileInternal(IStorageFile storageFile, Guid notificationId,
        string projectCode,
        string projectFolderPath)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Workbook? dataWorkbook = null;
        Sheets? classesSheets = null;
        Sheets? dataSheets = null;
        Worksheet? classesWorksheet = null;
        Worksheet? dataWorksheet = null;
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
            Worksheets = Workbook.Worksheets;

            // Open the survey classes and survey data file
            classesWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            dataWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Dati{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;
            dataSheets = dataWorkbook.Sheets;

            // For each worksheet in the input file, find the corresponding worksheet in the files and populate it
            foreach (Worksheet sheet in Worksheets)
            {
                // Find the corresponding worksheets in the app's files.
                // If there are none, create the sheets.
                try
                {
                    classesWorksheet = classesSheets.Item[sheet.Name];
                }
                catch
                {
                    classesWorksheet = classesSheets.Add();
                    classesWorksheet.Name = sheet.Name;
                }

                try
                {
                    dataWorksheet = dataSheets.Item[sheet.Name];
                }
                catch
                {
                    dataWorksheet = dataSheets.Add();
                    dataWorksheet.Name = sheet.Name;
                }

                responseTableRange = sheet.UsedRange;

                var originalDataTable = responseTableRange.MakeDataTable();

                // Step 1: Make the questions column
                var newClassesDataTable = AddQuestionsColumn(originalDataTable);

                // Step 2: Add the field name column
                newClassesDataTable = AddFieldNameColumn(newClassesDataTable);

                // Step 3: Add the category column
                newClassesDataTable = AddCategoryColumn(newClassesDataTable);

                // Step 4: Write the new datatable to the classes worksheet
                newClassesDataTable.WriteToWorksheet(classesWorksheet, "Classi");

                // Step 5: Keep only the question data from the original datatable
                var newDataDataTable = KeepDataColumns(originalDataTable);

                // Step 6: Write the new datatable to the data worksheet
                newDataDataTable.WriteToWorksheet(dataWorksheet, "Dati");

                // Release the resources on each iteration
                Marshal.ReleaseComObject(responseTableRange);
                responseTableRange = null;
                Marshal.ReleaseComObject(dataWorksheet);
                dataWorksheet = null;
                Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                Marshal.ReleaseComObject(sheet);
            }

            classesWorkbook.Save();
            dataWorkbook.Save();

            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();

            notificationService.ShowNotification("Importazione completata",
                "Il file è stato importato con successo.");
        }
        catch (Exception e)
        {
            notificationService.ShowNotification("Importazione fallita",
                "Si è verificato un errore durante l'importazione del file: " + e.Message);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);

            if (dataWorksheet is not null) Marshal.ReleaseComObject(dataWorksheet);
            if (classesWorksheet is not null) Marshal.ReleaseComObject(classesWorksheet);

            if (dataSheets is not null) Marshal.ReleaseComObject(dataSheets);
            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (dataWorkbook is not null)
            {
                dataWorkbook.Close(false);
                Marshal.ReleaseComObject(dataWorkbook);
            }

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
            .Where(s => s.StartsWith('D') && !s.Contains("PUNTO DI CAMPIONAMENTO"))
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

    private static DataTable KeepDataColumns(DataTable oldDataTable)
    {
        // Create the new datatable with only the data columns
        var newDataTable = new DataTable();

        // Extract the data columns from the original datatable
        var dataColumns = oldDataTable.Columns
            .Cast<DataColumn>()
            .Where(c => c.ColumnName.StartsWith('D'))
            .ToArray();

        // Add the data columns to the new datatable
        foreach (var column in dataColumns)
            newDataTable.Columns.Add(new DataColumn(column.ColumnName, typeof(string)));

        // Populate the new datatable with the data from the original datatable
        foreach (DataRow row in oldDataTable.Rows)
        {
            var newRow = newDataTable.NewRow();
            foreach (var column in dataColumns)
                newRow[column.ColumnName] = row[column.ColumnName];
            newDataTable.Rows.Add(newRow);
        }

        return newDataTable;
    }

    #endregion

    #region ActiveViewing file import

    public async Task ImportActiveViewingFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, string productCode)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ImportActiveViewingFileInternal(storageFile, notificationId, projectCode, projectFolderPath,
                    productCode)));
    }

    private void ImportActiveViewingFileInternal(IStorageFile storageFile, Guid notificationId,
        string projectCode,
        string projectFolderPath, string productCode)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Workbook? dataWorkbook = null;
        Sheets? classesSheets = null;
        Sheets? dataSheets = null;
        Worksheet? classesWorksheet = null;
        Worksheet? dataWorksheet = null;
        Range? tableRange = null;

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
            Worksheets = Workbook.Sheets;

            // Open the survey classes and survey data file
            classesWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
            dataWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Dati{projectCode}.xlsx"));
            classesSheets = classesWorkbook.Sheets;
            dataSheets = dataWorkbook.Sheets;

            // Find ActiveViewing's classes file and import it
            foreach (Worksheet sheet in Worksheets)
            {
                if (sheet.Name == "Input")
                {
                    dataWorksheet = dataSheets.Item[productCode];

                    tableRange = sheet.UsedRange;

                    var originalDataTable = tableRange.MakeDataTable();

                    // Step 1: Filter the datatable data by product code
                    var newDataTable = GetTableByProductCode(originalDataTable, productCode);

                    // Step 2 : Write the new datatable to the data worksheet
                    newDataTable.WriteToWorksheet(dataWorksheet, "Dati");
                }
                else if (sheet.Name.Contains("Classi domande"))
                {
                    classesWorksheet = classesSheets.Item[productCode];

                    tableRange = sheet.UsedRange;

                    var originalDataTable = tableRange.MakeDataTable();

                    // Step 1: Remove the useless rows
                    var newDataTable = originalDataTable.RemoveLastRows(3);

                    // Step 2: Remove the useless columns
                    newDataTable = newDataTable.RemoveLastColumns(6);

                    // Step 3 : Take the needed columns and rename them
                    newDataTable = TakeColumnsAndRename(newDataTable);

                    // Step 4: Write the new datatable to the classes worksheet
                    newDataTable.WriteToWorksheet(classesWorksheet, "Classi");
                }

                // Release the resources on each iteration
                if (tableRange is not null) Marshal.ReleaseComObject(tableRange);
                tableRange = null;
                if (dataWorksheet is not null) Marshal.ReleaseComObject(dataWorksheet);
                dataWorksheet = null;
                if (classesWorksheet is not null) Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                Marshal.ReleaseComObject(sheet);
            }

            classesWorkbook.Save();
            dataWorkbook.Save();

            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();

            notificationService.ShowNotification("Importazione completata",
                "Il file è stato importato con successo.");
        }
        catch (Exception e)
        {
            notificationService.ShowNotification("Importazione fallita",
                "Si è verificato un errore durante l'importazione del file: " + e.Message);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (tableRange is not null) Marshal.ReleaseComObject(tableRange);

            if (dataWorksheet is not null) Marshal.ReleaseComObject(dataWorksheet);
            if (classesWorksheet is not null) Marshal.ReleaseComObject(classesWorksheet);

            if (dataSheets is not null) Marshal.ReleaseComObject(dataSheets);
            if (classesSheets is not null) Marshal.ReleaseComObject(classesSheets);

            if (dataWorkbook is not null)
            {
                dataWorkbook.Close(false);
                Marshal.ReleaseComObject(dataWorkbook);
            }

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
        var newColumns = oldDataTable
            .Clone()
            .Columns
            .Cast<DataColumn>()
            .Where((_, index) => index is 1 or 4 or 5)
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

        foreach (var column in newColumns.Reverse())
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

    private static DataTable GetTableByProductCode(DataTable dataTable, string productCode)
    {
        // Take only the rows related to the given product
        var newDataTable = dataTable.AsEnumerable()
            .Where(row => row.Field<string>("Prodotto") == productCode)
            .CopyToDataTable();

        // Take the columns to remove
        var columnsToRemove = newDataTable.Columns
            .Cast<DataColumn>()
            .Where(c => !c.ColumnName.StartsWith("D.") && c.ColumnName != "LegCampionamento")
            .ToArray();

        // Remove the useless columns
        foreach (var column in columnsToRemove) newDataTable.Columns.Remove(column);

        // Rename first column
        newDataTable.Columns[0].ColumnName = "D.1 PUNTO DI CAMPIONAMENTO";

        return newDataTable;
    }

    #endregion
}