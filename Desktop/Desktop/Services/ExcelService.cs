using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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
        Worksheet? evaluationsWorksheet = null;
        Worksheet? expectationsWorksheet = null;

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
                    evaluationsWorksheet = (Worksheet)Worksheets[1];
                else
                    // Add new sheet after the last one
                    evaluationsWorksheet = (Worksheet)Worksheets.Add();

                // Rename the worksheet to match the product code
                evaluationsWorksheet.Name = project.Products.ElementAt(i).Code;

                // Add new sheet after the last one
                expectationsWorksheet = (Worksheet)Worksheets.Add();

                // Rename the worksheet to match the product code
                expectationsWorksheet.Name = $"ASP_{project.Products.ElementAt(i).Code}";

                // Release the resources on each iteration
                Marshal.ReleaseComObject(evaluationsWorksheet);
                evaluationsWorksheet = null;
                Marshal.ReleaseComObject(expectationsWorksheet);
                expectationsWorksheet = null;
            }

            Workbook.SaveAs(excelFilePath);
        }
        finally
        {
            if (evaluationsWorksheet is not null) Marshal.ReleaseComObject(evaluationsWorksheet);
            if (expectationsWorksheet is not null) Marshal.ReleaseComObject(expectationsWorksheet);
        }
    }

    #endregion

    #region Penelope file import

    public async Task ImportPenelopeFileAsync(IStorageFile storageFile, Guid notificationId, string projectCode,
        string projectFolderPath, ICollection<Product> products)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ImportPenelopeFileInternal(storageFile, notificationId, projectCode, projectFolderPath, products)));
    }

    private void ImportPenelopeFileInternal(IStorageFile storageFile, Guid notificationId,
        string projectCode,
        string projectFolderPath,
        ICollection<Product> products)
    {
        // Track the COM classes to be released
        Workbook? classesWorkbook = null;
        Workbook? dataWorkbook = null;
        Sheets? classesSheets = null;
        Sheets? dataSheets = null;
        Worksheet? classesWorksheet = null;
        Worksheet? evaluationsWorksheet = null;
        Worksheet? expectationsWorksheet = null;
        ListObjects? tables = null;
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
                if (!sheet.Name.Contains("ASP", StringComparison.CurrentCultureIgnoreCase))
                {
                    // Find the corresponding worksheets in the app's files.
                    // If there are none, create the sheets.
                    try
                    {
                        classesWorksheet = classesSheets[sheet.Name];

                        // Clean the previous table if there is any
                        tables = classesWorksheet.ListObjects;

                        foreach (ListObject table in tables)
                        {
                            table.Delete();
                            Marshal.ReleaseComObject(table);
                        }

                        Marshal.ReleaseComObject(tables);
                        tables = null;
                    }
                    catch
                    {
                        classesWorksheet = classesSheets.Add();
                        classesWorksheet.Name = sheet.Name;
                    }

                    try
                    {
                        evaluationsWorksheet = dataSheets[sheet.Name];

                        // Clean the previous table if there is any
                        tables = evaluationsWorksheet.ListObjects;

                        foreach (ListObject table in tables)
                        {
                            table.Delete();
                            Marshal.ReleaseComObject(table);
                        }
                    }
                    catch
                    {
                        evaluationsWorksheet = dataSheets.Add();
                        evaluationsWorksheet.Name = sheet.Name;
                    }

                    responseTableRange = sheet.UsedRange;

                    var originalDataTable = responseTableRange.MakeDataTable();

                    // Step 1: Make the questions column
                    var newClassesDataTable = AddQuestionsColumn(originalDataTable,
                        products.First(product => product.Code == sheet.Name).DisplayName);

                    // Step 2: Add the field name column
                    newClassesDataTable = AddFieldNameColumn(newClassesDataTable);

                    // Step 3: Add the category column
                    newClassesDataTable = AddCategoryColumn(newClassesDataTable);

                    // Step 4: Write the new datatable to the classes worksheet
                    newClassesDataTable.WriteToWorksheet(classesWorksheet, "Classi");

                    // Step 5: Keep only the question data from the original datatable
                    var newDataDataTable = KeepDataColumns(originalDataTable);

                    // Step 6: Write the new datatable to the data worksheet
                    newDataDataTable.WriteToWorksheet(evaluationsWorksheet, "Dati");
                }

                // Release the resources on each iteration
                if (responseTableRange is not null) Marshal.ReleaseComObject(responseTableRange);
                responseTableRange = null;
                if (tables is not null) Marshal.ReleaseComObject(tables);
                tables = null;
                if (evaluationsWorksheet is not null) Marshal.ReleaseComObject(evaluationsWorksheet);
                evaluationsWorksheet = null;
                if (classesWorksheet is not null) Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                Marshal.ReleaseComObject(sheet);
            }

            // For the expectations sheet in the input file, partition the data and write it in the right sheet
            expectationsWorksheet = Worksheets["ASP"];

            responseTableRange = expectationsWorksheet.UsedRange;

            var expectationsDataTable = responseTableRange.MakeDataTable();

            Marshal.ReleaseComObject(expectationsWorksheet);
            expectationsWorksheet = null;
            Marshal.ReleaseComObject(responseTableRange);
            responseTableRange = null;

            var expectationsByCode = from dataRow in expectationsDataTable.AsEnumerable().AsQueryable()
                group dataRow by dataRow.Field<string?>("sigla")
                into codeGroup
                select codeGroup;

            foreach (var group in expectationsByCode)
            {
                // Find the corresponding worksheet in the app's files.
                // If there is none, create the sheet.
                try
                {
                    expectationsWorksheet = dataSheets[group.Key];

                    // Clean the previous table if there is any
                    tables = expectationsWorksheet.ListObjects;

                    foreach (ListObject table in tables)
                    {
                        table.Delete();
                        Marshal.ReleaseComObject(table);
                    }
                }
                catch
                {
                    expectationsWorksheet = dataSheets.Add();
                    expectationsWorksheet.Name = group.Key;
                }

                var expectationsTable = group.CopyToDataTable();
                expectationsTable.Columns.Remove("NumQuestionario");
                expectationsTable.Columns.Remove("sigla");
                expectationsTable.WriteToWorksheet(expectationsWorksheet, "Aspettative");

                // Release the resources on each iteration
                if (tables is not null) Marshal.ReleaseComObject(tables);
                tables = null;
                Marshal.ReleaseComObject(expectationsWorksheet);
                expectationsWorksheet = null;
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

            if (tables is not null) Marshal.ReleaseComObject(tables);

            if (expectationsWorksheet is not null) Marshal.ReleaseComObject(expectationsWorksheet);
            if (evaluationsWorksheet is not null) Marshal.ReleaseComObject(evaluationsWorksheet);
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

    private static DataTable AddQuestionsColumn(DataTable oldDataTable, string productDisplayName)
    {
        // Create the new datatable with the questions column
        var newDataTable = new DataTable();
        newDataTable.Columns.Add(new DataColumn(productDisplayName, typeof(string)));

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

    #region Product file creation

    public async Task CreateProductFilesAsync(Guid notificationId, ICollection<Product> products,
        string projectFolderPath, string projectCode)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                CreateProductFilesInternal(notificationId, products, projectFolderPath, projectCode)));
    }

    private void CreateProductFilesInternal(Guid notificationId, ICollection<Product> products,
        string projectFolderPath,
        string projectCode)
    {
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
        Worksheet? evaluationsWorksheet = null;
        Worksheet? expectationsWorksheet = null;

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

                try
                {
                    classesWorksheet = classesSheets[product.Code];
                    evaluationsWorksheet = dataSheets[product.Code];
                    expectationsWorksheet = dataSheets[$"ASP_{product.Code}"];
                }
                catch
                {
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
                        (uint)productCurrentCount,
                        (uint)--productTotalCount).GetAwaiter().GetResult();

                    continue;
                }

                // Copy the sheets and paste them in the product file and rename
                classesWorksheet.Copy(After: worksheet);
                Marshal.ReleaseComObject(classesWorksheet);
                classesWorksheet = null;
                classesWorksheet = worksheets[worksheets.Count];
                classesWorksheet.Name = "Classi";

                evaluationsWorksheet.Copy(After: classesWorksheet);
                Marshal.ReleaseComObject(evaluationsWorksheet);
                evaluationsWorksheet = null;
                evaluationsWorksheet = worksheets[worksheets.Count];
                evaluationsWorksheet.Name = "Dati";

                expectationsWorksheet.Copy(After: evaluationsWorksheet);
                Marshal.ReleaseComObject(expectationsWorksheet);
                expectationsWorksheet = null;
                expectationsWorksheet = worksheets[worksheets.Count];
                expectationsWorksheet.Name = "Aspettative";

                // Delete the first empty sheet in the product file
                worksheet.Delete();

                var excelFilePath =
                    Path.Combine(Path.Combine(projectFolderPath, "Elaborazioni"), $"{product.DisplayName.Trim()}.xlsx");
                workbook.SaveAs(excelFilePath);

                // Release the resources on each iteration
                Marshal.ReleaseComObject(expectationsWorksheet);
                expectationsWorksheet = null;
                Marshal.ReleaseComObject(evaluationsWorksheet);
                evaluationsWorksheet = null;
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
            if (expectationsWorksheet is not null) Marshal.ReleaseComObject(expectationsWorksheet);
            if (evaluationsWorksheet is not null) Marshal.ReleaseComObject(evaluationsWorksheet);
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

    public enum SynopticTableType
    {
        Confezione,
        GradimentoComplessivo,
        SoddisfazioneComplessiva,
        PropensioneAlRiconsumo,
        ConfrontoProdottoAbituale
    }

    public async Task ProcessProductFilesAsync(Guid notificationId, ICollection<string> fileNames)
    {
        await Task.Run(() =>
            ExecuteWithCleanup(() =>
                ProcessProductFileInternal(notificationId, fileNames)));
    }

    private void ProcessProductFileInternal(Guid notificationId, ICollection<string> fileNames)
    {
        // Create a silent Excel application
        ExcelApp = new Application
        {
            Visible = false,
            DisplayAlerts = false
        };
        Workbooks = ExcelApp.Workbooks;

        var processedFiles = 0u;

        Workbook? workbook = null;

        try
        {
            foreach (var fileName in fileNames)
            {
                workbook = Workbooks.Open(fileName);

                ProcessClosedTable(workbook, TableType.Scale9, Path.GetFileNameWithoutExtension(fileName));
                ProcessClosedTable(workbook, TableType.Scale5, Path.GetFileNameWithoutExtension(fileName));
                _ = ProcessFrequencyTable(workbook, TableType.Scale9, Path.GetFileNameWithoutExtension(fileName));
                var scale5FrequencyTables = ProcessFrequencyTable(workbook, TableType.Scale5,
                    Path.GetFileNameWithoutExtension(fileName));
                ProcessAdequacyTable(workbook, scale5FrequencyTables);
                ProcessSynopticTable(workbook);

                workbook.Save();

                notificationService.UpdateProgressNotificationAsync(notificationId,
                    "Elaborazione file prodotti in corso...",
                    ++processedFiles,
                    (uint)fileNames.Count).GetAwaiter().GetResult();

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

    private static void ProcessClosedTable(Workbook workbook, TableType scale, string productDisplayName)
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
                destinationSheet = worksheets[scale == TableType.Scale5 ? "Tabelle 5" : "Tabelle 9"];

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
                let classe = classRow.Field<string?>("Classe") ?? ""
                where (scale == TableType.Scale5 &&
                       Regex.IsMatch(classe, "[AI]", RegexOptions.IgnoreCase | RegexOptions.Compiled,
                           TimeSpan.FromMilliseconds(100))) ||
                      (scale == TableType.Scale9 &&
                       string.Compare(classe, "G", StringComparison.CurrentCultureIgnoreCase) == 0)
                select new
                {
                    Question = classRow.Field<string?>(productDisplayName),
                    Label = classRow.Field<string?>("Etichetta")
                };

            // If the table is 9-scaled, put the "Gradimento complessivo" / "Soddisfazione complessiva" label at the first position
            if (scale == TableType.Scale9)
            {
                var firstRow = questionsAndLabels.FirstOrDefault(q =>
                    string.Compare(q.Label, "Gradimento complessivo", StringComparison.CurrentCultureIgnoreCase) == 0 ||
                    string.Compare(q.Label, "Soddisfazione complessiva", StringComparison.CurrentCultureIgnoreCase) ==
                    0);

                if (firstRow is not null)
                    questionsAndLabels = questionsAndLabels
                        .Where(q =>
                            string.Compare(q.Label, "Gradimento complessivo",
                                StringComparison.CurrentCultureIgnoreCase) != 0 &&
                            string.Compare(q.Label, "Soddisfazione complessiva",
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

            ICollection<KeyValuePair<string, DataTable>> dataTables = [];

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
                        .Select(row => Convert.ToDouble(row.Field<string?>(qAndL.Question.Trim())))
                        .Where(value => scale != TableType.Scale5 || Convert.ToUInt32(value) != 6)
                        .ToList();

                    var average = values.Average();
                    var lsd = 1.96 * (Math.Sqrt(values.Select(v => Math.Pow(v - average, 2)).Sum() / values.Count) /
                                      Math.Sqrt(values.Count));

                    var newRow = genericTable.NewRow();
                    newRow["Generale"] = string.IsNullOrEmpty(qAndL.Label.Trim())
                        ? qAndL.Question.Trim()
                        : qAndL.Label.Trim();
                    newRow["Media"] = average;
                    newRow["LSD"] = lsd;
                    genericTable.Rows.Add(newRow);
                }

                dataTables.Add(new KeyValuePair<string, DataTable>("Generale", genericTable));
            }

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
                        .Select(row => Convert.ToDouble(row.Field<string?>(qAndL.Question.Trim())))
                        .Where(value => scale != TableType.Scale5 || Convert.ToUInt32(value) != 6)
                        .ToList();

                    var average = values.Average();
                    var lsd = 1.96 * (Math.Sqrt(values.Select(v => Math.Pow(v - average, 2)).Sum() / values.Count) /
                                      Math.Sqrt(values.Count));

                    var newRow = locationTable.NewRow();
                    newRow[location.ApplyCase(LetterCasing.Sentence)] =
                        string.IsNullOrEmpty(qAndL.Label.Trim()) ? qAndL.Question.Trim() : qAndL.Label.Trim();
                    newRow["Media"] = average;
                    newRow["LSD"] = lsd;
                    locationTable.Rows.Add(newRow);
                }

                dataTables.Add(new KeyValuePair<string, DataTable>(location, locationTable));
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

    private static IEnumerable<KeyValuePair<(string, char), DataTable>> ProcessFrequencyTable(Workbook workbook,
        TableType scale, string productDisplayName)
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
                destinationSheet = worksheets[scale == TableType.Scale5 ? "Frequenze 5" : "Frequenze 9"];

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
                let classe = classRow.Field<string?>("Classe") ?? ""
                where (scale == TableType.Scale5 &&
                       Regex.IsMatch(classe, "[AI]", RegexOptions.IgnoreCase | RegexOptions.Compiled,
                           TimeSpan.FromMilliseconds(100))) ||
                      (scale == TableType.Scale9 &&
                       string.Compare(classe, "G", StringComparison.CurrentCultureIgnoreCase) == 0)
                select new
                {
                    Question = classRow.Field<string?>(productDisplayName),
                    Label = classRow.Field<string?>("Etichetta"),
                    Class = Convert.ToChar(classe.Trim())
                };

            // Put the "Gradimento complessivo" or "Soddisfazione complessiva" label at the first position if scale is 9
            if (scale == TableType.Scale9)
                questionsAndLabels = questionsAndLabels.OrderByDescending(q =>
                    string.Compare(q.Label, "Gradimento complessivo", StringComparison.CurrentCultureIgnoreCase) == 0 ||
                    string.Compare(q.Label, "Soddisfazione complessiva", StringComparison.CurrentCultureIgnoreCase) ==
                    0);

            // Find the columns to delete
            var columnsToRemove = dataDataTable.Columns
                .Cast<DataColumn>()
                .Where(column => column.ColumnName != "D.1 PUNTO DI CAMPIONAMENTO" &&
                                 !questionsAndLabels.Any(q => string.Compare(q.Question, column.ColumnName,
                                     StringComparison.CurrentCultureIgnoreCase) == 0))
                .ToList();

            foreach (var column in columnsToRemove) dataDataTable.Columns.Remove(column);

            ICollection<KeyValuePair<string, DataTable>> dataTables = [];
            ICollection<KeyValuePair<(string, char), DataTable>> cumulativeDataTables = [];

            // For each question/label, create the frequency table
            foreach (var qAndL in questionsAndLabels)
            {
                // Compute the number of people who answered with a certain value and the percentage, excluding the invalid ones
                var excludedCount = dataDataTable.AsEnumerable()
                    .Count(dataRow =>
                    {
                        var value = Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question.Trim()));
                        return (TableType.Scale5 == scale && value is < 1 or > 5) ||
                               (TableType.Scale9 == scale && value is < 1 or > 9);
                    });

                var results = from dataRow in dataDataTable.AsEnumerable().AsQueryable()
                    orderby Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question.Trim()))
                    group dataRow by Convert.ToUInt32(dataRow.Field<string?>(qAndL.Question.Trim()))
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
                frequencyTable.Columns.Add(new DataColumn(qAndL.Label.Trim(), typeof(uint)));
                frequencyTable.Columns.Add(new DataColumn("Percentuale", typeof(double)));
                frequencyTable.Columns.Add(new DataColumn("Totale", typeof(uint)));

                foreach (var result in results)
                {
                    var newRow = frequencyTable.NewRow();
                    newRow[qAndL.Label.Trim()] = Convert.ToUInt32(result.Value);
                    newRow["Percentuale"] = result.Percentage;
                    newRow["Totale"] = result.Count;

                    frequencyTable.Rows.Add(newRow);
                }

                dataTables.Add(new KeyValuePair<string, DataTable>(qAndL.Label.Trim(), frequencyTable));

                // Compute the cumulative results
                var cumulativeFrequencyTable = new DataTable();
                cumulativeFrequencyTable.Columns.Add(new DataColumn(qAndL.Label.Trim(), typeof(string)));
                cumulativeFrequencyTable.Columns.Add(new DataColumn("Percentuale", typeof(double)));
                cumulativeFrequencyTable.Columns.Add(new DataColumn("Totale", typeof(uint)));

                if (scale == TableType.Scale5)
                {
                    // Group results into partitions and add rows for each partition
                    // Do not add the "da 1 a 2" row for the "Confronto abituale" label
                    if (string.Compare(qAndL.Label.Trim(), "Confronto abituale",
                            StringComparison.CurrentCultureIgnoreCase) != 0)
                    {
                        var partition1 = results.Where(r => r.Value == 1 || r.Value == 2);
                        var row1 = cumulativeFrequencyTable.NewRow();
                        row1[qAndL.Label.Trim()] = "da 1 a 2";
                        row1["Percentuale"] = partition1.Sum(r => r.Percentage);
                        row1["Totale"] = partition1.Sum(r => r.Count);
                        cumulativeFrequencyTable.Rows.Add(row1);
                    }

                    // Add the "3" and "4" rows instead of the "3" and "da 4 a 5" rows for the "Propensione al riconsumo" label
                    if (string.Compare(qAndL.Label.Trim(), "Propensione al riconsumo",
                            StringComparison.CurrentCultureIgnoreCase) != 0)
                    {
                        var partition2 = results.Where(r => r.Value == 3);
                        var partition3 = results.Where(r => r.Value == 4 || r.Value == 5);

                        var row2 = cumulativeFrequencyTable.NewRow();
                        row2[qAndL.Label.Trim()] = "3";
                        row2["Percentuale"] = partition2.Sum(r => r.Percentage);
                        row2["Totale"] = partition2.Sum(r => r.Count);
                        cumulativeFrequencyTable.Rows.Add(row2);

                        var row3 = cumulativeFrequencyTable.NewRow();
                        row3[qAndL.Label.Trim()] = "da 4 a 5";
                        row3["Percentuale"] = partition3.Sum(r => r.Percentage);
                        row3["Totale"] = partition3.Sum(r => r.Count);
                        cumulativeFrequencyTable.Rows.Add(row3);
                    }
                    else
                    {
                        var partition2 = results.Where(r => r.Value == 4);
                        var partition3 = results.Where(r => r.Value == 5);

                        var row2 = cumulativeFrequencyTable.NewRow();
                        row2[qAndL.Label.Trim()] = "4";
                        row2["Percentuale"] = partition2.Sum(r => r.Percentage);
                        row2["Totale"] = partition2.Sum(r => r.Count);
                        cumulativeFrequencyTable.Rows.Add(row2);

                        var row3 = cumulativeFrequencyTable.NewRow();
                        row3[qAndL.Label.Trim()] = "5";
                        row3["Percentuale"] = partition3.Sum(r => r.Percentage);
                        row3["Totale"] = partition3.Sum(r => r.Count);
                        cumulativeFrequencyTable.Rows.Add(row3);
                    }
                }
                else
                {
                    // Group results into three partitions
                    var partition1 = results.Where(r => r.Value >= 1 && r.Value <= 3);
                    var partition2 = results.Where(r => r.Value >= 4 && r.Value <= 6);
                    var partition3 = results.Where(r => r.Value >= 7 && r.Value <= 9);

                    // Add rows for each partition
                    var row1 = cumulativeFrequencyTable.NewRow();
                    row1[qAndL.Label.Trim()] = "da 1 a 3";
                    row1["Percentuale"] = partition1.Sum(r => r.Percentage);
                    row1["Totale"] = partition1.Sum(r => r.Count);
                    cumulativeFrequencyTable.Rows.Add(row1);

                    var row2 = cumulativeFrequencyTable.NewRow();
                    row2[qAndL.Label.Trim()] = "da 4 a 6";
                    row2["Percentuale"] = partition2.Sum(r => r.Percentage);
                    row2["Totale"] = partition2.Sum(r => r.Count);
                    cumulativeFrequencyTable.Rows.Add(row2);

                    var row3 = cumulativeFrequencyTable.NewRow();
                    row3[qAndL.Label.Trim()] = "da 7 a 9";
                    row3["Percentuale"] = partition3.Sum(r => r.Percentage);
                    row3["Totale"] = partition3.Sum(r => r.Count);
                    cumulativeFrequencyTable.Rows.Add(row3);
                }

                cumulativeDataTables.Add(
                    new KeyValuePair<(string, char), DataTable>((qAndL.Label.Trim(), qAndL.Class),
                        cumulativeFrequencyTable));
            }

            // Write all the datatables to the worksheet
            foreach (var kvp in dataTables) kvp.Value.WriteFrequencyTableToWorksheet(destinationSheet, kvp.Key);
            foreach (var kvp in cumulativeDataTables)
                kvp.Value.WriteCumulativeFrequencyTableToWorksheet(destinationSheet, kvp.Key.Item1);

            return cumulativeDataTables;
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

    private static void ProcessAdequacyTable(Workbook workbook,
        IEnumerable<KeyValuePair<(string, char), DataTable>> scale5FrequencyTables)
    {
        Sheets? worksheets = null;
        Worksheet? lastSheet = null;
        Worksheet? destinationSheet = null;
        ListObjects? tables = null;

        try
        {
            worksheets = workbook.Worksheets;
            lastSheet = worksheets[worksheets.Count];

            // Check if the sheet already exists, if so clean it
            try
            {
                destinationSheet = worksheets["Adeguatezze"];

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
                destinationSheet = worksheets.Add(After: lastSheet);
                destinationSheet.Name = "Adeguatezze";
            }

            var scale5FrequencyTablesArray = scale5FrequencyTables.ToArray();

            var transposedAdequacyTables = scale5FrequencyTablesArray.Where(kvp => kvp.Key.Item2 is 'a' or 'A')
                .Select(kvp => new KeyValuePair<string, DataTable>(kvp.Key.Item1, kvp.Value.Transpose()));

            var adequacyTable = new DataTable();
            adequacyTable.Columns.Add("Attributo", typeof(string));
            adequacyTable.Columns.Add("Troppo poco", typeof(double));
            adequacyTable.Columns.Add("Giusto", typeof(double));
            adequacyTable.Columns.Add("Troppo", typeof(double));

            foreach (var kvp in transposedAdequacyTables)
            {
                // Do not insert the "Propensione al riconsumo" and "Confronto abituale" attributes
                if (Regex.IsMatch(kvp.Key, "Propensione al riconsumo|Confronto abituale",
                        RegexOptions.IgnoreCase | RegexOptions.Compiled,
                        TimeSpan.FromMilliseconds(100))) continue;

                var row = adequacyTable.NewRow();
                row["Attributo"] = kvp.Key;
                row["Troppo poco"] = (double)kvp.Value.Rows[0]["da 1 a 2"] * 100;
                row["Giusto"] = (double)kvp.Value.Rows[0]["3"] * 100;
                row["Troppo"] = (double)kvp.Value.Rows[0]["da 4 a 5"] * 100;
                adequacyTable.Rows.Add(row);
            }

            adequacyTable.WriteAdequacyTableToWorksheet(destinationSheet, "Adeguatezze");

            var transposedIntensityTables = scale5FrequencyTablesArray.Where(kvp => kvp.Key.Item2 is 'i' or 'I')
                .Select(kvp => new KeyValuePair<string, DataTable>(kvp.Key.Item1, kvp.Value.Transpose()));

            var intensityTable = new DataTable();
            intensityTable.Columns.Add("Attributo", typeof(string));
            intensityTable.Columns.Add("Per niente", typeof(double));
            intensityTable.Columns.Add("Abbastanza", typeof(double));
            intensityTable.Columns.Add("Estremamente", typeof(double));

            foreach (var kvp in transposedIntensityTables)
            {
                var row = intensityTable.NewRow();
                row["Attributo"] = kvp.Key;
                row["Per niente"] = (double)kvp.Value.Rows[0]["da 1 a 2"] * 100;
                row["Abbastanza"] = (double)kvp.Value.Rows[0]["3"] * 100;
                row["Estremamente"] = (double)kvp.Value.Rows[0]["da 4 a 5"] * 100;
                intensityTable.Rows.Add(row);
            }

            intensityTable.WriteIntensityTableToWorksheet(destinationSheet, "Intensità");
        }
        finally // Clean up the resources not managed by the base class
        {
            if (tables is not null) Marshal.ReleaseComObject(tables);
            if (destinationSheet is not null) Marshal.ReleaseComObject(destinationSheet);
            if (lastSheet is not null) Marshal.ReleaseComObject(lastSheet);
            if (worksheets is not null) Marshal.ReleaseComObject(worksheets);
        }
    }

    private static void ProcessSynopticTable(Workbook workbook)
    {
        Sheets? worksheets = null;
        Worksheet? lastSheet = null;
        Worksheet? destinationSheet = null;
        Worksheet? sourceSheet = null;
        ListObjects? tables = null;
        Range? dataRange = null;

        try
        {
            worksheets = workbook.Worksheets;
            lastSheet = worksheets[worksheets.Count];

            // Check if the sheet already exists, if so clean it
            try
            {
                destinationSheet = worksheets["Sinottiche"];

                // Clean the previous table if there is any
                tables = destinationSheet.ListObjects;

                foreach (ListObject t in tables)
                {
                    t.Delete();
                    Marshal.ReleaseComObject(t);
                }

                tables = null;
            }
            catch
            {
                destinationSheet = worksheets.Add(After: lastSheet);
                destinationSheet.Name = "Sinottiche";
            }

            #region Confezione

            // Create the synoptic table structure
            var synopticTable = new DataTable();
            synopticTable.Columns.Add("Macrocategoria");
            synopticTable.Columns.Add("Categoria");
            synopticTable.Columns.Add("Attributo", typeof(string));
            synopticTable.Columns.Add("Valore", typeof(double));

            // Read the "Aspettative" table from Aspettative to get overall rating
            sourceSheet = worksheets["Aspettative"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("Aspettative", StringComparison.CurrentCultureIgnoreCase))
                    dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            var expectationsDataTable = dataRange?.MakeDataTable();

            if (expectationsDataTable is null) return;

            for (var i = 1; i < 4; i++)
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Aspettativa";
                newRow["Categoria"] = "Confezione";
                newRow["Attributo"] = i switch
                {
                    1 => "Gradimento atteso",
                    2 => "Fiducia",
                    3 => "Gradimento confezione",
                    _ => ""
                };
                newRow["Valore"] = expectationsDataTable.AsEnumerable().Average(row => Convert.ToDouble(row[i]));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            synopticTable.WriteSynopticTableToWorksheet(destinationSheet, "Confezione", SynopticTableType.Confezione);

            #endregion

            #region Gradimento complessivo / Soddifazione complessiva + profilo di gradimento

            // Create the synoptic table structure
            synopticTable = new DataTable();
            synopticTable.Columns.Add("Macrocategoria");
            synopticTable.Columns.Add("Categoria");
            synopticTable.Columns.Add("Attributo", typeof(string));
            synopticTable.Columns.Add("Valore", typeof(double));

            // Read the "Generale" table from Tabelle 9 to get overall rating
            sourceSheet = worksheets["Tabelle 9"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("Generale", StringComparison.CurrentCultureIgnoreCase)) dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            var overallDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().ToArray();

            if (overallDataRows is null) return;

            foreach (var overallDataRow in overallDataRows.Take(1))
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = overallDataRow.Field<string?>("Generale");
                newRow["Attributo"] = overallDataRow.Field<string?>("Generale");
                newRow["Valore"] = Convert.ToDouble(overallDataRow.Field<string?>("Media"));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            // Read the "C_Gradimento complessivo" or "C_Soddisfazione complessiva" table from Frequenze 9 to get cumulative rating
            sourceSheet = worksheets["Frequenze 9"];
            tables = sourceSheet.ListObjects;
            var synopticTableType = SynopticTableType.GradimentoComplessivo;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("C_Gradimento complessivo", StringComparison.CurrentCultureIgnoreCase))
                {
                    dataRange = table.Range;
                }
                else if (table.Name.Contains("C_Soddisfazione complessiva", StringComparison.CurrentCultureIgnoreCase))
                {
                    dataRange = table.Range;
                    synopticTableType = SynopticTableType.SoddisfazioneComplessiva;
                }

                Marshal.ReleaseComObject(table);
            }

            var overallFrequencyDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().SkipLast(1);

            if (overallFrequencyDataRows is null) return;

            foreach (var overallFrequencyDataRow in overallFrequencyDataRows)
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = synopticTableType == SynopticTableType.GradimentoComplessivo
                    ? "Gradimento complessivo"
                    : "Soddisfazione complessiva";
                newRow["Attributo"] = synopticTableType == SynopticTableType.GradimentoComplessivo
                    ? overallFrequencyDataRow.Field<string?>("Gradimento complessivo")
                    : overallFrequencyDataRow.Field<string?>("Soddisfazione complessiva");
                newRow["Valore"] = Convert.ToDouble(overallFrequencyDataRow.Field<string?>("Percentuale"));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            foreach (var overallDataRow in overallDataRows.Skip(1))
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = "Profilo di gradimento";
                newRow["Attributo"] = overallDataRow.Field<string?>("Generale");
                newRow["Valore"] = Convert.ToDouble(overallDataRow.Field<string?>("Media"));
                synopticTable.Rows.Add(newRow);
            }

            synopticTable.WriteSynopticTableToWorksheet(
                destinationSheet,
                synopticTableType == SynopticTableType.GradimentoComplessivo
                    ? "Gradimento complessivo"
                    : "Soddisfazione complessiva",
                synopticTableType);

            #endregion

            #region Propensione al riconsumo

            // Create the synoptic table structure
            synopticTable = new DataTable();
            synopticTable.Columns.Add("Macrocategoria");
            synopticTable.Columns.Add("Categoria");
            synopticTable.Columns.Add("Attributo", typeof(string));
            synopticTable.Columns.Add("Valore", typeof(double));

            // Read the "Generale" table from Tabelle 5 to get overall rating
            sourceSheet = worksheets["Tabelle 5"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("Generale", StringComparison.CurrentCultureIgnoreCase)) dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            overallDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().ToArray();

            if (overallDataRows is null) return;

            foreach (var overallDataRow in overallDataRows.Where(dataRow =>
                         string.Compare(dataRow["Generale"].ToString(), "Propensione al riconsumo",
                             StringComparison.CurrentCultureIgnoreCase) == 0))
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = "Propensione al riconsumo";
                newRow["Attributo"] = overallDataRow.Field<string?>("Generale");
                newRow["Valore"] = Convert.ToDouble(overallDataRow.Field<string?>("Media"));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            // Read the "C_Propensione al riconsumo" table from Frequenze 5 to get cumulative rating
            sourceSheet = worksheets["Frequenze 5"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("C_Propensione al riconsumo", StringComparison.CurrentCultureIgnoreCase))
                    dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            overallFrequencyDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().SkipLast(1);

            if (overallFrequencyDataRows is null) return;

            foreach (var overallFrequencyDataRow in overallFrequencyDataRows)
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = "Propensione al riconsumo";
                newRow["Attributo"] = overallFrequencyDataRow.Field<string?>("Propensione al riconsumo");
                newRow["Valore"] = Convert.ToDouble(overallFrequencyDataRow.Field<string?>("Percentuale"));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            synopticTable.WriteSynopticTableToWorksheet(destinationSheet, "Propensione al riconsumo",
                SynopticTableType.PropensioneAlRiconsumo);

            #endregion

            #region Confronto abituale

            // Create the synoptic table structure
            synopticTable = new DataTable();
            synopticTable.Columns.Add("Macrocategoria");
            synopticTable.Columns.Add("Categoria");
            synopticTable.Columns.Add("Attributo", typeof(string));
            synopticTable.Columns.Add("Valore", typeof(double));

            // Read the "Generale" table from Tabelle 5 to get overall rating
            sourceSheet = worksheets["Tabelle 5"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("Generale", StringComparison.CurrentCultureIgnoreCase)) dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            overallDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().ToArray();

            if (overallDataRows is null) return;

            foreach (var overallDataRow in overallDataRows.Where(dataRow =>
                         string.Compare(dataRow["Generale"].ToString(), "Confronto abituale",
                             StringComparison.CurrentCultureIgnoreCase) == 0))
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = "Confronto abituale";
                newRow["Attributo"] = overallDataRow.Field<string?>("Generale");
                newRow["Valore"] = Convert.ToDouble(overallDataRow.Field<string?>("Media"));
                synopticTable.Rows.Add(newRow);
            }

            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            dataRange = null;
            Marshal.ReleaseComObject(tables);
            tables = null;
            Marshal.ReleaseComObject(sourceSheet);
            sourceSheet = null;

            // Read the "C_Confronto abituale" table from Frequenze 5 to get cumulative rating
            sourceSheet = worksheets["Frequenze 5"];
            tables = sourceSheet.ListObjects;

            foreach (ListObject table in tables)
            {
                if (table.Name.Contains("C_Confronto abituale", StringComparison.CurrentCultureIgnoreCase))
                    dataRange = table.Range;

                Marshal.ReleaseComObject(table);
            }

            overallFrequencyDataRows = dataRange?.MakeDataTable().Rows.Cast<DataRow>().SkipLast(1);

            if (overallFrequencyDataRows is null) return;

            foreach (var overallFrequencyDataRow in overallFrequencyDataRows)
            {
                var newRow = synopticTable.NewRow();
                newRow["Macrocategoria"] = "Valutazione";
                newRow["Categoria"] = "Confronto abituale";
                newRow["Attributo"] = overallFrequencyDataRow.Field<string?>("Confronto abituale");
                newRow["Valore"] = Convert.ToDouble(overallFrequencyDataRow.Field<string?>("Percentuale"));
                synopticTable.Rows.Add(newRow);
            }

            synopticTable.WriteSynopticTableToWorksheet(destinationSheet, "Confronto abituale",
                SynopticTableType.ConfrontoProdottoAbituale);

            #endregion
        }
        finally
        {
            if (dataRange is not null) Marshal.ReleaseComObject(dataRange);
            if (tables is not null) Marshal.ReleaseComObject(tables);
            if (sourceSheet is not null) Marshal.ReleaseComObject(sourceSheet);
            if (destinationSheet is not null) Marshal.ReleaseComObject(destinationSheet);
            if (lastSheet is not null) Marshal.ReleaseComObject(lastSheet);
            if (worksheets is not null) Marshal.ReleaseComObject(worksheets);
        }
    }

    #endregion
}