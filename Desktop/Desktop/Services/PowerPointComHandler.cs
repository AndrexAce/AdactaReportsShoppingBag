using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal abstract class PowerPointComHandler : BaseComHandler
{
    protected Application? PowerPointApp;
    protected Presentations? Presentations;

    protected override void ReleaseComObjects()
    {
        // Release COM objects to prevent memory leaks
        if (Presentations is not null) Marshal.ReleaseComObject(Presentations);
        Presentations = null;

        if (PowerPointApp is not null)
        {
            PowerPointApp.Quit();
            Marshal.ReleaseComObject(PowerPointApp);
        }

        PowerPointApp = null;
    }
}