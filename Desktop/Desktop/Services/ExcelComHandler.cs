using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

[SuppressMessage("Critical Code Smell", "S1215:\"GC.Collect\" should not be called",
    Justification = "COM objects lifetime should be manually managed.")]
internal abstract class ExcelComHandler : BaseComHandler
{
    protected Application? ExcelApp;
    protected Workbook? Workbook;
    protected Workbooks? Workbooks;
    protected Sheets? Worksheets;

    protected override void ReleaseComObjects()
    {
        // Release COM objects to prevent memory leaks
        if (Worksheets is not null) Marshal.ReleaseComObject(Worksheets);
        Worksheets = null;

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