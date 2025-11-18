using System;
using System.Threading.Tasks;
using Microsoft.Windows.AppNotifications;
using Microsoft.Windows.AppNotifications.Builder;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal class NotificationService : INotificationService
{
    public Guid ShowProgressNotification(string title, string message)
    {
        var tag = Guid.NewGuid();

        var notification = new AppNotificationBuilder()
            .AddText(title)
            .AddText(message)
            .BuildNotification();

        notification.Tag = tag.ToString();
        notification.ExpiresOnReboot = false;

        AppNotificationManager.Default.Show(notification);

        // Return the tag so caller can delete the notification later
        return tag;
    }

    public async Task RemoveNotification(Guid tag)
    {
        await AppNotificationManager.Default.RemoveByTagAsync(tag.ToString()).AsTask();
    }
}