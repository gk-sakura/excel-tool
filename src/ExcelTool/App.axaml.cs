using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using ExcelTool.Services.FileSystem;
using ExcelTool.Services.Notifications;
using ExcelTool.Services.Workflows;
using ExcelTool.ViewModels;
using ExcelTool.Views;

namespace ExcelTool;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var mainWindow = new MainWindow();
            var filePickerService = new FilePickerService(mainWindow.StorageProvider);
            mainWindow.DataContext = new MainViewModel(
                new MultiExcelMergeService(),
                new MultiSheetMergeService(),
                new CreateFolderFromExcelService(),
                filePickerService,
                new NotificationService(mainWindow));
            desktop.MainWindow = mainWindow;
        }

        base.OnFrameworkInitializationCompleted();
    }
}