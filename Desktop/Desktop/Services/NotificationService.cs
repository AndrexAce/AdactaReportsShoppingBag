using System;
using System.Threading.Tasks;
using Microsoft.Windows.AppNotifications;
using Microsoft.Windows.AppNotifications.Builder;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal class NotificationService : INotificationService
{
    public Guid ShowNotification(string title, string message)
    {
        var tag = Guid.NewGuid();

        var notification = new AppNotificationBuilder()
            .AddText(title)
            .AddText(message)
            .BuildNotification();

        notification.Tag = tag.ToString();
        notification.Group = tag.ToString();
        notification.ExpiresOnReboot = true;

        AppNotificationManager.Default.Show(notification);

        // Return the tag so caller can delete the notification later
        return tag;
    }

    public async Task<Guid> ShowProgressNotificationAsync(string title, string message, string statusMessage, uint totalToProcess)
    {
        var tag = Guid.NewGuid();

        var notification = new AppNotificationBuilder()
            .AddText(title)
            .AddText(message)
            .AddProgressBar(new AppNotificationProgressBar()
            .BindStatus()
            .BindValue()
            .BindValueStringOverride())
            .BuildNotification();

        notification.Tag = tag.ToString();
        notification.Group = tag.ToString();
        notification.ExpiresOnReboot = true;

        AppNotificationManager.Default.Show(notification);

        await UpdateProgressNotificationAsync(tag, statusMessage, 0, totalToProcess);

        // Return the tag so caller can delete or update the notification later
        return tag;
    }

    public async Task RemoveNotificationAsync(Guid tag)
    {
        await AppNotificationManager.Default.RemoveByTagAndGroupAsync(tag.ToString(), tag.ToString());
    }

    public async Task UpdateProgressNotificationAsync(Guid tag, string statusMessage, uint processed, uint totalToProcess)
    {
        var data = new AppNotificationProgressData(processed + 1)
        {
            Status = statusMessage,
            Value = (double) processed / totalToProcess,
            ValueStringOverride = $"{processed}/{totalToProcess}"
        };

        await AppNotificationManager.Default.UpdateAsync(data, tag.ToString(), tag.ToString());
    }
}