using System;
using System.Threading;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal abstract class BaseComHandler
{
    protected void ExecuteWithCleanup(Action operation)
    {
        try
        {
            operation();
        }
        finally
        {
            ReleaseComObjects();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }

    protected abstract void ReleaseComObjects();
}