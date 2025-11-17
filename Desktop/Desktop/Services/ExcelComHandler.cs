using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

[SuppressMessage("Critical Code Smell", "S1215:\"GC.Collect\" should not be called", Justification = "COM objects lifetime should be manually managed.")]
internal class ExcelComHandler : BaseComHandler
{
    protected Application? excelApp = null;
    protected Workbooks? workbooks = null;
    protected Workbook? workbook = null;
    protected Sheets? sheets = null;
    protected Collection<Worksheet> worksheets = [];

    override public void ReleaseCOMObjects()
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
        sheets = null;

        if (workbook is not null)
        {
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
        }
        workbook = null;

        if (workbooks is not null)
        {
            workbooks.Close();
            Marshal.ReleaseComObject(workbooks);
        }
        workbooks = null;

        if (excelApp is not null)
        {
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
        excelApp = null;

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
