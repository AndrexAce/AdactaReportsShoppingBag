using System;
using System.Threading.Tasks;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface INotificationService
{
    public Guid ShowNotification(string title, string message);

    public Task RemoveNotificationAsync(Guid tag);
}