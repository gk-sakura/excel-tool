namespace ExcelTool.Services.Notifications;

public interface INotificationService
{
    void ShowInfo(string title, string message);
    void ShowError(string title, string message);
    void ShowSuccess(string title, string message);
    void ShowWarning(string title, string message);
}