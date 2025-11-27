using System;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface INotificationService
{
    public Guid ShowNotification(string title, string message);

    public Task<Guid> ShowProgressNotificationAsync(string title, string message, string statusMessage,
        uint totalToProcess);

    public Task RemoveNotificationAsync(Guid tag);

    public Task UpdateProgressNotificationAsync(Guid tag, string statusMessage, uint processed, uint totalToProcess);
}