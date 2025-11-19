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
        }
    }

    protected async Task ExecuteWithCleanupAsync(Func<Task> operation)
    {
        // Capture the current synchronization context to ensure COM cleanup happens on the correct thread
        var syncContext = SynchronizationContext.Current;

        try
        {
            await operation();
        }
        finally
        {
            if (syncContext != null)
            {
                syncContext.Post(_ => ReleaseComObjects(), null);
            }
            else
            {
                ReleaseComObjects();
            }
        }
    }

    protected abstract void ReleaseComObjects();
}