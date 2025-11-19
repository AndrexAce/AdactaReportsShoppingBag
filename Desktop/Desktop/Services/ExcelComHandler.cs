using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

[SuppressMessage("Critical Code Smell", "S1215:\"GC.Collect\" should not be called",
    Justification = "COM objects lifetime should be manually managed.")]
internal abstract class ExcelComHandler : BaseComHandler
{
    protected readonly Collection<Worksheet> Worksheets = [];
    protected Application? ExcelApp;
    protected Sheets? Sheets;
    protected Workbook? Workbook;
    protected Workbooks? Workbooks;

    protected override void ReleaseComObjects()
    {
        // Release COM objects to prevent memory leaks
        foreach (var element in Worksheets) Marshal.ReleaseComObject(element);
        Worksheets.Clear();

        if (Sheets is not null) Marshal.ReleaseComObject(Sheets);
        Sheets = null;

        if (Workbook is not null)
        {
            Workbook.Close(false);
            Marshal.ReleaseComObject(Workbook);
        }

        Workbook = null;

        if (Workbooks is not null) Marshal.ReleaseComObject(Workbooks);
        Workbooks = null;

        if (ExcelApp is not null)
        {
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
        }

        ExcelApp = null;
    }
}