using System;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal abstract class BaseComHandler
{
    public void ExecuteWithCleanup(Action operation)
    {
        try
        {
            operation();
        }
        finally
        {
            ReleaseCOMObjects();
        }
    }

    public async Task ExecuteWithCleanupAsync(Func<Task> operation)
    {
        try
        {
            await operation();
        }
        finally
        {
            ReleaseCOMObjects();
        }
    }

    public abstract void ReleaseCOMObjects();
}
