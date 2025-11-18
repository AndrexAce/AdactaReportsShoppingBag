using System;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface INotificationService
{
    public Guid ShowProgressNotification(string title, string message);

    public Task RemoveNotification(Guid tag);
}