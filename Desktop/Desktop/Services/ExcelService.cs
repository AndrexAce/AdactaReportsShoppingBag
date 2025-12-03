using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Storage;
using AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;
using AdactaInternational.AdactaReportsShoppingBag.Model;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using Humanizer;
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
        ListObjects? tables = null;
        ListObject? table = null;
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

                    // Clean the previous table if there is any
                    tables = classesWorksheet.ListObjects;

                    if (tables.Count > 0)
                    {
                        table = tables["Classi"];
                        table.Delete();
                    }
                }
                catch
                {
                    classesWorksheet = classesSheets.Add();
                    classesWorksheet.Name = sheet.Name;
                }

                try
                {
                    dataWorksheet = dataSheets.Item[sheet.Name];

                    // Clean the previous table if there is any
                    tables = dataWorksheet.ListObjects;

                    if (tables.Count > 0)
                    {
                        table = tables["Dati"];
                        table.Delete();
                    }
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
                if (table is not null) Marshal.ReleaseComObject(table);
                table = null;
                if (tables is not null) Marshal.ReleaseComObject(tables);
                tables = null;
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

            if (table is not null) Marshal.ReleaseComObject(table);

            if (tables is not null) Marshal.ReleaseComObject(tables);

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
            .Where(s => s.StartsWith("D", StringComparison.CurrentCultureIgnoreCase) &&
                        !s.Contains("PUNTO DI CAMPIONAMENTO", StringComparison.CurrentCultureIgnoreCase))
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
        ListObjects? tables = null;
        ListObject? table = null;
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
                if (string.Compare(sheet.Name, "Input", StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    // Find the corresponding worksheet in the app's files.
                    // If there is none, create the sheet.
                    try
                    {
                        dataWorksheet = dataSheets.Item[productCode];

                        // Clean the previous table if there is any
                        tables = dataWorksheet.ListObjects;

                        if (tables.Count > 0)
                        {
                            table = tables["Dati"];
                            table.Delete();
                        }
                    }
                    catch
                    {
                        dataWorksheet = dataSheets.Add();
                        dataWorksheet.Name = productCode;
                    }

                    tableRange = sheet.UsedRange;

                    var originalDataTable = tableRange.MakeDataTable();

                    // Step 1: Filter the datatable data by product code
                    var newDataTable = GetTableByProductCode(originalDataTable, productCode);

                    // Step 2 : Write the new datatable to the data worksheet
                    newDataTable.WriteToWorksheet(dataWorksheet, "Dati");
                }
                else if (sheet.Name.Contains("Classi domande", StringComparison.CurrentCultureIgnoreCase))
                {
                    // Find the corresponding worksheet in the app's files.
                    // If there is none, create the sheet.
                    try
                    {
                        classesWorksheet = classesSheets.Item[productCode];

                        // Clean the previous table if there is any
                        tables = classesWorksheet.ListObjects;

                        if (tables.Count > 0)
                        {
                            table = tables["Classi"];
                            table.Delete();
                        }
                    }
                    catch
                    {
                        classesWorksheet = classesSheets.Add();
                        classesWorksheet.Name = productCode;
                    }

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
                if (table is not null) Marshal.ReleaseComObject(table);
                table = null;
                if (tables is not null) Marshal.ReleaseComObject(tables);
                tables = null;
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

            if (table is not null) Marshal.ReleaseComObject(table);

            if (tables is not null) Marshal.ReleaseComObject(tables);

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
                if (string.Compare(c.ColumnName, "Testo Domanda", StringComparison.CurrentCultureIgnoreCase) == 0)
                    c.ColumnName = "Domanda";
                else if (string.Compare(c.ColumnName, "Etichetta Domanda", StringComparison.CurrentCultureIgnoreCase) ==
                         0) c.ColumnName = "Etichetta";

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
            .Where(row =>
                string.Compare(row.Field<string?>("Prodotto"), productCode,
                    StringComparison.CurrentCultureIgnoreCase) ==
                0 ||
                string.Compare(row.Field<string?>("prodotto"), productCode,
                    StringComparison.CurrentCultureIgnoreCase) ==
                0)
            .CopyToDataTable();

        // Take the columns to remove
        var columnsToRemove = newDataTable.Columns
            .Cast<DataColumn>()
            .Where(c => !c.ColumnName.StartsWith("D.", StringComparison.CurrentCultureIgnoreCase) &&
                        string.Compare(c.ColumnName, "LegCampionamento", StringComparison.CurrentCultureIgnoreCase) !=
                        0)
            .ToArray();

        // Remove the useless columns
        foreach (var column in columnsToRemove) newDataTable.Columns.Remove(column);

        // Rename first column
        newDataTable.Columns[0].ColumnName = "D.1 PUNTO DI CAMPIONAMENTO";

        return newDataTable;
    }

    #endregion

    #region Product file creation

    public async Task CreateProductFilesAsync(Guid notificationId, ICollection<Product> products,
        string projectFolderPath, string projectCode)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ProcessExcelFilesInternal(notificationId, products, projectFolderPath, projectCode)));
    }

    private void ProcessExcelFilesInternal(Guid notificationId, ICollection<Product> products, string projectFolderPath,
        string projectCode)
    {
        // Filter the products to create files only for those that don't have one yet
        var productTotalCount = products.Count;
        var productCurrentCount = 0;

        // Track the COM classes to be released
        Workbook? workbook = null;
        Sheets? worksheets = null;
        Worksheet? worksheet = null;
        Workbook? classesWorkbook = null;
        Workbook? dataWorkbook = null;
        Sheets? classesSheets = null;
        Sheets? dataSheets = null;
        Worksheet? classesWorksheet = null;
        Worksheet? dataWorksheet = null;

        try
        {
            if (!Directory.Exists(Path.Combine(projectFolderPath, "Elaborazioni")))
                Directory.CreateDirectory(Path.Combine(projectFolderPath, "Elaborazioni"));

            // Create a silent Excel application
            ExcelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            Workbooks = ExcelApp.Workbooks;

            foreach (var product in products)
            {
                workbook = Workbooks.Add();
                worksheets = workbook.Worksheets;
                worksheet = worksheets[1];

                // Open the survey classes and survey data file
                classesWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Classi{projectCode}.xlsx"));
                dataWorkbook = Workbooks.Open(Path.Combine(projectFolderPath, $"Dati{projectCode}.xlsx"));
                classesSheets = classesWorkbook.Sheets;
                dataSheets = dataWorkbook.Sheets;
                classesWorksheet = classesSheets.Item[product.Code];
                dataWorksheet = dataSheets.Item[product.Code];

                // Copy the sheets and paste them in the product file and rename
                classesWorksheet.Copy(After: worksheet);
                Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                classesWorksheet = worksheets[worksheets.Count];
                classesWorksheet.Name = "Classi";

                dataWorksheet.Copy(After: classesWorksheet);
                Marshal.ReleaseComObject(dataWorksheet);
                dataWorksheet = null;
                dataWorksheet = worksheets[worksheets.Count];
                dataWorksheet.Name = "Dati";

                // Delete the first empty sheet in the product file
                worksheet.Delete();

                var excelFilePath =
                    Path.Combine(Path.Combine(projectFolderPath, "Elaborazioni"), $"{product.DisplayName.Trim()}.xlsx");
                workbook.SaveAs(excelFilePath);

                // Release the resources on each iteration
                Marshal.ReleaseComObject(dataWorksheet);
                dataWorksheet = null;
                Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                Marshal.ReleaseComObject(dataSheets);
                dataSheets = null;
                Marshal.ReleaseComObject(classesSheets);
                classesSheets = null;
                dataWorkbook.Close(false);
                Marshal.ReleaseComObject(dataWorkbook);
                dataWorkbook = null;
                classesWorkbook.Close(false);
                Marshal.ReleaseComObject(classesWorkbook);
                classesWorkbook = null;
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
                Marshal.ReleaseComObject(worksheets);
                worksheets = null;
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                workbook = null;

                notificationService.UpdateProgressNotificationAsync(notificationId,
                    "Creazione file prodotti in corso...",
                    (uint)++productCurrentCount,
                    (uint)productTotalCount).GetAwaiter().GetResult();
            }

            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();
        }
        catch (Exception e)
        {
            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();
            notificationService.ShowNotification("Elaborazione fallita",
                "Si è verificato un errore durante la creazione dei file di prodotti: " + e.Message);
        }
        finally // Clean up the resources not managed by the base class
        {
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

            if (worksheet is not null) Marshal.ReleaseComObject(worksheet);

            if (worksheets is not null) Marshal.ReleaseComObject(worksheets);

            if (workbook is not null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
        }
    }

    #endregion

    #region Product file processing

    private enum TableType
    {
        Scale5,
        Scale9
    }

    public async Task ProcessProductFilesAsync(Guid notificationId, string projectFolderPath)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ProcessProductFileInternal(notificationId, projectFolderPath)));
    }

    private void ProcessProductFileInternal(Guid notificationId, string projectFolderPath)
    {
        // Create a silent Excel application
        ExcelApp = new Application
        {
            Visible = false,
            DisplayAlerts = false
        };
        Workbooks = ExcelApp.Workbooks;

        var fileNames = Directory.GetFiles(Path.Combine(projectFolderPath, "Elaborazioni"), "*.xlsx");
        var processedFiles = 0u;

        Workbook? workbook = null;

        try
        {
            foreach (var fileName in fileNames)
            {
                workbook = Workbooks.Open(fileName);

                ProcessClosedTable(workbook, TableType.Scale9);
                ProcessClosedTable(workbook, TableType.Scale5);
                ProcessFrequenciesTable(workbook, TableType.Scale9);
                ProcessFrequenciesTable(workbook, TableType.Scale5);

                workbook.Save();

                notificationService.UpdateProgressNotificationAsync(notificationId,
                    "Elaborazione file prodotti in corso...",
                    ++processedFiles,
                    (uint)fileNames.Length).GetAwaiter().GetResult();

                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                workbook = null;
            }

            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();
            notificationService.ShowNotification("Elaborazione completata",
                "I file dei prodotti sono stati elaborati con successo.");
        }
        catch (Exception e)
        {
            notificationService.RemoveNotificationAsync(notificationId).GetAwaiter().GetResult();
            notificationService.ShowNotification("Elaborazione fallita",
                "Si è verificato un errore durante l'elaborazione dei file dei prodotti: " + e.Message);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (workbook is not null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
        }
    }

    private static void ProcessClosedTable(Workbook workbook, TableType scale)
    {
        if (scale != TableType.Scale5 && scale != TableType.Scale9)
            throw new ArgumentOutOfRangeException(nameof(scale), "Invalid table type.");

        Sheets? worksheets = null;
        Worksheet? classesSheet = null;
        Worksheet? dataSheet = null;
        Worksheet? destinationSheet = null;
        ListObjects? tables = null;
        Range? dataRange = null;

        try
        {
            worksheets = workbook.Worksheets;
            classesSheet = worksheets["Classi"];
            dataSheet = worksheets["Dati"];

            // Check if the sheet already exists, if so clean it
            try
            {
                destinationSheet = worksheets.Item[scale == TableType.Scale5 ? "Tabelle 5" : "Tabelle 9"];

                // Clean the previous table if there is any
                tables = destinationSheet.ListObjects;

                foreach (ListObject table in tables)
                {
                    table.Delete();
                    Marshal.ReleaseComObject(table);
                }
            }
            catch
            {
                destinationSheet = worksheets.Add(After: dataSheet);
                destinationSheet.Name = scale == TableType.Scale5 ? "Tabelle 5" : "Tabelle 9";
            }

            dataRange = classesSheet.UsedRange;
            var classesDataTable = dataRange.MakeDataTable();
            Marshal.ReleaseComObject(dataRange);
            dataRange = null;

            dataRange = dataSheet.UsedRange;
            var dataDataTable = dataRange.MakeDataTable();
            Marshal.ReleaseComObject(dataRange);
            dataRange = null;

            // Take the questions and labels
            var questionsAndLabels = from classRow in classesDataTable.AsEnumerable().AsQueryable()
                let classe = classRow.Field<string?>("Classe")
                where classe == (scale == TableType.Scale5 ? "A" : "G") ||
                      classe == (scale == TableType.Scale5 ? "a" : "g")
                select new
                {
                    Question = classRow.Field<string?>("Domanda"),
                    Label = classRow.Field<string?>("Etichetta")
                };

            // If the table is 9-scaled, put the "Gradimento complessivo" label at the first position
            if (scale == TableType.Scale9)
            {
                var firstRow = questionsAndLabels.FirstOrDefault(q =>
                    string.Compare(q.Label, "Gradimento complessivo", StringComparison.CurrentCultureIgnoreCase) == 0);

                if (firstRow is not null)
                    questionsAndLabels = questionsAndLabels
                        .Where(q => string.Compare(q.Label, "Gradimento complessivo",
                            StringComparison.CurrentCultureIgnoreCase) != 0)
                        .Prepend(firstRow);
            }

            // Find the columns to delete
            var columnsToRemove = dataDataTable.Columns
                .Cast<DataColumn>()
                .Where(column => column.ColumnName != "D.1 PUNTO DI CAMPIONAMENTO" &&
                                 !questionsAndLabels.Any(q => string.Compare(q.Question, column.ColumnName,
                                     StringComparison.CurrentCultureIgnoreCase) == 0))
                .ToList();

            foreach (var column in columnsToRemove) dataDataTable.Columns.Remove(column);

            // Take the list of possible locations and create groups
            var locations = from dataRow in dataDataTable.AsEnumerable().AsQueryable()
                group dataRow by dataRow.Field<string?>("D.1 PUNTO DI CAMPIONAMENTO")
                into locationGroup
                select locationGroup.Key;

            IEnumerable<KeyValuePair<string, DataTable>> dataTables = [];

            // Create a table for each location
            foreach (var location in locations)
            {
                var locationTable = new DataTable();
                locationTable.Columns.Add(new DataColumn(location.ApplyCase(LetterCasing.Sentence), typeof(string)));
                // Add the product name column
                locationTable.Columns.Add(new DataColumn("Media", typeof(double)));
                // Add the lsd column
                locationTable.Columns.Add(new DataColumn("LSD", typeof(double)));

                // For each question, calculate the average and lsd
                foreach (var qAndL in questionsAndLabels)
                {
                    var questionRows = dataDataTable.AsEnumerable()
                        .Where(row => string.Compare(row.Field<string?>("D.1 PUNTO DI CAMPIONAMENTO"), location,
                            StringComparison.CurrentCultureIgnoreCase) == 0);

                    var values = questionRows
                        .Select(row => Convert.ToDouble(row.Field<string?>(qAndL.Question)))
                        .ToList();

                    var average = values.Average();
                    var lsd = 1.96 * (Math.Sqrt(values.Select(v => Math.Pow(v - average, 2)).Sum() / values.Count) /
                                      Math.Sqrt(values.Count));

                    var newRow = locationTable.NewRow();
                    newRow[location.ApplyCase(LetterCasing.Sentence)] =
                        string.IsNullOrEmpty(qAndL.Label) ? qAndL.Question : qAndL.Label;
                    newRow["Media"] = average;
                    newRow["LSD"] = lsd;
                    locationTable.Rows.Add(newRow);
                }

                dataTables = dataTables.Append(new KeyValuePair<string, DataTable>(location, locationTable));
            }

            if (locations.Any())
            {
                // Create the generic table
                var genericTable = new DataTable();
                genericTable.Columns.Add(new DataColumn("Generale", typeof(string)));
                // Add the product name column
                genericTable.Columns.Add(new DataColumn("Media", typeof(double)));
                // Add the lsd column
                genericTable.Columns.Add(new DataColumn("LSD", typeof(double)));

                // For each question, calculate the average and lsd
                foreach (var qAndL in questionsAndLabels)
                {
                    var questionRows = dataDataTable.AsEnumerable();

                    var values = questionRows
                        .Select(row => Convert.ToDouble(row.Field<string?>(qAndL.Question)))
                        .ToList();

                    var average = values.Average();
                    var lsd = 1.96 * (Math.Sqrt(values.Select(v => Math.Pow(v - average, 2)).Sum() / values.Count) /
                                      Math.Sqrt(values.Count));

                    var newRow = genericTable.NewRow();
                    newRow["Generale"] = string.IsNullOrEmpty(qAndL.Label) ? qAndL.Question : qAndL.Label;
                    newRow["Media"] = average;
                    newRow["LSD"] = lsd;
                    genericTable.Rows.Add(newRow);
                }

                dataTables = dataTables.Prepend(new KeyValuePair<string, DataTable>("Generale", genericTable));
            }

            // Write all the datatables to the worksheet
            foreach (var kvp in dataTables) kvp.Value.WriteClosedTableToWorksheet(destinationSheet, kvp.Key);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            if (tables is not null) Marshal.ReleaseComObject(tables);
            if (destinationSheet is not null) Marshal.ReleaseComObject(destinationSheet);
            if (dataSheet is not null) Marshal.ReleaseComObject(dataSheet);
            if (classesSheet is not null) Marshal.ReleaseComObject(classesSheet);
            if (worksheets is not null) Marshal.ReleaseComObject(worksheets);
        }
    }

    private static void ProcessFrequenciesTable(Workbook workbook, TableType scale)
    {
        if (scale != TableType.Scale5 && scale != TableType.Scale9)
            throw new ArgumentOutOfRangeException(nameof(scale), "Invalid table type.");

        Sheets? worksheets = null;
        Worksheet? classesSheet = null;
        Worksheet? dataSheet = null;
        Worksheet? destinationSheet = null;
        ListObjects? tables = null;
        Range? dataRange = null;

        try
        {
            worksheets = workbook.Worksheets;
            classesSheet = worksheets["Classi"];
            dataSheet = worksheets["Dati"];

            // Check if the sheet already exists, if so clean it
            try
            {
                destinationSheet = worksheets.Item[scale == TableType.Scale5 ? "Frequenze 5" : "Frequenze 9"];

                // Clean the previous table if there is any
                tables = destinationSheet.ListObjects;

                foreach (ListObject table in tables)
                {
                    table.Delete();
                    Marshal.ReleaseComObject(table);
                }
            }
            catch
            {
                destinationSheet = worksheets.Add(After: dataSheet);
                destinationSheet.Name = scale == TableType.Scale5 ? "Frequenze 5" : "Frequenze 9";
            }

            dataRange = classesSheet.UsedRange;
            var classesDataTable = dataRange.MakeDataTable();
            Marshal.ReleaseComObject(dataRange);
            dataRange = null;

            dataRange = dataSheet.UsedRange;
            var dataDataTable = dataRange.MakeDataTable();
            Marshal.ReleaseComObject(dataRange);
            dataRange = null;

            // Take the questions and labels
            var questionsAndLabels = from classRow in classesDataTable.AsEnumerable().AsQueryable()
                let classe = classRow.Field<string?>("Classe")
                where classe == (scale == TableType.Scale5 ? "A" : "G") ||
                      classe == (scale == TableType.Scale5 ? "a" : "g")
                select new
                {
                    Question = classRow.Field<string?>("Domanda"),
                    Label = classRow.Field<string?>("Etichetta")
                };

            // Find the columns to delete
            var columnsToRemove = dataDataTable.Columns
                .Cast<DataColumn>()
                .Where(column => column.ColumnName != "D.1 PUNTO DI CAMPIONAMENTO" &&
                                 !questionsAndLabels.Any(q => string.Compare(q.Question, column.ColumnName,
                                     StringComparison.CurrentCultureIgnoreCase) == 0))
                .ToList();

            foreach (var column in columnsToRemove) dataDataTable.Columns.Remove(column);

            IEnumerable<KeyValuePair<string, DataTable>> dataTables = [];

            // For each question/label, create the frequency table
            foreach (var qAndL in questionsAndLabels)
            {
                // Compute the number of people who answered with a certain value and the percentage, excluding the invalid ones
                var excludedCount = dataDataTable.AsEnumerable()
                    .Count(dataRow =>
                    {
                        var value = Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question));
                        return (TableType.Scale5 == scale && value is < 1 or > 5) ||
                               (TableType.Scale9 == scale && value is < 1 or > 9);
                    });


                var results = from dataRow in dataDataTable.AsEnumerable().AsQueryable()
                    orderby Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question))
                    group dataRow by Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question))
                    into resultGroup
                    where (TableType.Scale5 == scale && resultGroup.Key >= 1 && resultGroup.Key <= 5) ||
                          (TableType.Scale9 == scale && resultGroup.Key >= 1 && resultGroup.Key <= 9)
                    select new
                    {
                        Value = resultGroup.Key,
                        Percentage = (double)resultGroup.Count() / (dataDataTable.Rows.Count - excludedCount),
                        Count = resultGroup.Count()
                    };

                // If some data is missing, add it
                if ((scale == TableType.Scale5 && results.Count() != 5) ||
                    (scale == TableType.Scale9 && results.Count() != 9))
                {
                    // Find out the missing values
                    var values = Enumerable.Sequence(1, scale == TableType.Scale5 ? 5 : 9, 1);
                    var missingValues = values.Except(results.Select(r => Convert.ToInt32(r.Value)));

                    // Add them with 0 count and 0 percentage
                    results = missingValues.Aggregate(results,
                        (current, missingValue) =>
                            current.Append(new { Value = (uint)missingValue, Percentage = 0.0, Count = 0 }));
                    // Sort the results again
                    results = results.OrderBy(r => r.Value);
                }

                // Create the table that contains the data and add it
                var frequencyTable = new DataTable();
                frequencyTable.Columns.Add(new DataColumn(qAndL.Label, typeof(uint)));
                frequencyTable.Columns.Add(new DataColumn("Percentuale", typeof(double)));
                frequencyTable.Columns.Add(new DataColumn("Totale", typeof(uint)));

                foreach (var result in results)
                {
                    var newRow = frequencyTable.NewRow();
                    newRow[qAndL.Label] = Convert.ToUInt32(result.Value);
                    newRow["Percentuale"] = result.Percentage;
                    newRow["Totale"] = result.Count;

                    frequencyTable.Rows.Add(newRow);
                }

                dataTables = dataTables.Append(new KeyValuePair<string, DataTable>(qAndL.Label, frequencyTable));
            }

            // Write all the datatables to the worksheet
            foreach (var kvp in dataTables) kvp.Value.WriteFrequenciesTableToWorksheet(destinationSheet, kvp.Key);
        }
        finally // Clean up the resources not managed by the base class
        {
            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            if (tables is not null) Marshal.ReleaseComObject(tables);
            if (destinationSheet is not null) Marshal.ReleaseComObject(destinationSheet);
            if (dataSheet is not null) Marshal.ReleaseComObject(dataSheet);
            if (classesSheet is not null) Marshal.ReleaseComObject(classesSheet);
            if (worksheets is not null) Marshal.ReleaseComObject(worksheets);
        }
    }

    #endregion
}