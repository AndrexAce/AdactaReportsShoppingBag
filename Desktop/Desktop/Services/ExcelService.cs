using AdactaInternational.AdactaReportsShoppingBag.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class ExcelService : IExcelService
{
    [SuppressMessage("Critical Code Smell", "S1215:\"GC.Collect\" should not be called", Justification = "COM objects lifetime should be manually managed.")]
    public void CreateExcelClassesFile(ReportPrj project, string projectFolderPath)
    {
        // Keep references to COM objects to release them later
        Application? excelApp = null;
        Workbooks? workbooks = null;
        Workbook? workbook = null;
        Sheets? sheets = null;
        Collection<Worksheet> worksheets = [];

        try
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
        finally
        {
            // Release COM objects to prevent memory leaks
            foreach (var element in worksheets)
            {
                Marshal.ReleaseComObject(element);
            }
            worksheets.Clear();

            if (sheets is not null)
            {
                Marshal.ReleaseComObject(sheets);
            }
            if (workbook is not null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            if (workbooks is not null)
            {
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
            }
            if (excelApp is not null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
