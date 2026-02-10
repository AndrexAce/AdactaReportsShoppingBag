using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal abstract class PowerPointComHandler : BaseComHandler
{
    protected Application? PowerPointApp;
    protected Presentation? Presentation;
    protected Presentations? Presentations;
    protected Slides? Slides;

    protected override void ReleaseComObjects()
    {
        // Release COM objects to prevent memory leaks
        if (Slides is not null) Marshal.ReleaseComObject(Slides);
        Slides = null;

        if (Presentation is not null)
        {
            Presentation.Close();
            Marshal.ReleaseComObject(Presentation);
        }

        Presentation = null;

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