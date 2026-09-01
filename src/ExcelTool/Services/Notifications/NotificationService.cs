using System;
using Avalonia.Controls;
using Avalonia.Controls.Notifications;

namespace ExcelTool.Services.Notifications;

public class NotificationService: INotificationService
{
    private readonly WindowNotificationManager _manager;

    public NotificationService(Window window)
    {
        _manager = new WindowNotificationManager(window)
        {
            Position = NotificationPosition.TopRight,
            MaxItems = 3
        };
    }

    public void ShowSuccess(string title, string message)
    {
        Show(title, message, NotificationType.Success);
    }

    public void ShowWarning(string title, string message)
    {
        Show(title, message, NotificationType.Warning);
    }

    public void ShowError(string title, string message)
    {
        Show(title, message, NotificationType.Error);
    }

    public void ShowInfo(string title, string message)
    {
        Show(title, message, NotificationType.Information);
    }

    private void Show(
        string title,
        string message,
        NotificationType type)
    {
        _manager.Show(
            new Notification(
                title,
                message,
                type,
                TimeSpan.FromSeconds(3)));
    }
}