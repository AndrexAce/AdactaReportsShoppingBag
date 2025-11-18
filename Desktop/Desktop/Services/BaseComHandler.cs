using System;
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
        try
        {
            await operation();
        }
        finally
        {
            ReleaseComObjects();
        }
    }

    protected abstract void ReleaseComObjects();
}