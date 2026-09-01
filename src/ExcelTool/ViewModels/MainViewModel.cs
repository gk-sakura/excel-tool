using System;
using System.Collections.Generic;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelTool.Services.FileSystem;
using ExcelTool.Services.Workflows;
using System.Threading.Tasks;
using ExcelTool.Services.Notifications;
using ExcelTool.Models;

namespace ExcelTool.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private readonly MultiExcelMergeService _multiExcelMergeService;
    private readonly MultiSheetMergeService _multiSheetMergeService;
    private readonly CreateFolderFromExcelService _createFolderFromExcelService;
    private readonly FilePickerService _filePickerService;
    private readonly INotificationService _notificationService;
    private readonly IProgress<double> _progressReporter;

    public IReadOnlyList<FunctionOption> Functions { get; } =
    [
        new(ExcelFunction.MergeWorkbooks, "合并工作簿"),
        new(ExcelFunction.MergeWorksheets, "合并工作表"),
        new(ExcelFunction.CreateFolders, "批量创建文件夹"),
    ];

    [ObservableProperty] private double _progress;
    [ObservableProperty] private string _sourceFolderPath = "";
    [ObservableProperty] private string _targetFolderPath = "";
    [ObservableProperty] private FunctionOption _selectedFunction;
    [ObservableProperty] private string _sourceFilePath = "";

    public bool ShowSelectSourceFolder => SelectedFunction.Value is ExcelFunction.MergeWorkbooks or ExcelFunction.MergeWorksheets;
    public bool ShowSelectTargetFolder => 
        SelectedFunction.Value is ExcelFunction.MergeWorkbooks or ExcelFunction.MergeWorksheets or ExcelFunction.CreateFolders;
    public bool ShowSelectSourceFile => SelectedFunction.Value is ExcelFunction.CreateFolders;

    partial void OnSelectedFunctionChanged(FunctionOption value)
    {
        OnPropertyChanged(nameof(ShowSelectSourceFolder));
        OnPropertyChanged(nameof(ShowSelectTargetFolder));
        OnPropertyChanged(nameof(ShowSelectSourceFile));
    }
    
    
    public MainViewModel(
        MultiExcelMergeService multiExcelMergeService,
        MultiSheetMergeService multiSheetMergeService,
        CreateFolderFromExcelService createFolderFromExcelService,
        FilePickerService filePickerService,
        INotificationService notificationService)
    {
        _multiExcelMergeService = multiExcelMergeService;
        _multiSheetMergeService = multiSheetMergeService;
        _createFolderFromExcelService = createFolderFromExcelService;
        _filePickerService = filePickerService;
        _notificationService = notificationService;
        
        _selectedFunction = Functions[0];
        _progressReporter = new Progress<double>(value => Progress = value);
        TargetFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    }

    /// <summary>
    /// 选择源文件路径
    /// </summary>
    [RelayCommand]
    private async Task SelectSourceFolderAsync()
    {
        var folderPaths = await _filePickerService.PickFolderAsync();
        if (folderPaths.Length == 0)
        {
            return;
        }
        SourceFolderPath = folderPaths[0];
    }

    /// <summary>
    /// 选择保存路径
    /// </summary>
    [RelayCommand]
    private async Task SelectTargetFolderAsync()
    {
        var folderPaths = await _filePickerService.PickFolderAsync();
        if (folderPaths.Length == 0)
        {
            return;
        }
        TargetFolderPath = folderPaths[0];
    }

    [RelayCommand]
    private async Task SelectSourceFileAsync()
    {
        var filePaths = await _filePickerService.PickFileAsync(["*.xlsx"]);
        if (filePaths.Length == 0)
        {
            return;
        }
        SourceFilePath = filePaths[0];
        var folderPath = Path.GetDirectoryName(filePaths[0]);
        if (folderPath is null)
        { 
            return;
        }
        TargetFolderPath = folderPath;
    }

    /// <summary>
    /// 执行按钮命令
    /// </summary>
    [RelayCommand]
    private async Task ExecuteCommandAsync()
    {
        if (ShowSelectSourceFolder && string.IsNullOrWhiteSpace(SourceFolderPath))
        {
            _notificationService.ShowWarning("提示", "请选择文件路径");
            return;
        }

        if (ShowSelectTargetFolder && string.IsNullOrWhiteSpace(TargetFolderPath))
        {
            _notificationService.ShowWarning("提示", "请选择保存路径");
            return;
        }

        if (ShowSelectSourceFile && string.IsNullOrWhiteSpace(SourceFilePath))
        {
            _notificationService.ShowWarning("提示", "请选择文件路径");
            return;
        }

        switch (SelectedFunction.Value)
        {
            case ExcelFunction.MergeWorkbooks:
                await MultiExcelMergeAsync();
                break;
            case ExcelFunction.MergeWorksheets:
                await MultiSheetMergeAsync();
                break;
            case ExcelFunction.CreateFolders:
                await CreateFolderFromExcelAsync();
                break;
            default:
                _notificationService.ShowWarning("提示", "请选择需要执行的功能");
                break;
        }
    }
    
    /// <summary>
    /// 合并多个工作簿中的工作表到一个工作表中
    /// </summary>
    private async Task MultiExcelMergeAsync()
    {
        var filePaths = await _filePickerService.GetFolderFilesAsync(SourceFolderPath, [".xlsx"]);
        if (filePaths.Length == 0)
        {
            return;
        }
        Progress = 0;
        using var workbook = await Task.Run(() => _multiExcelMergeService.Merge(filePaths, _progressReporter));
        await _filePickerService.SaveExcelAsync(
            workbook,
            TargetFolderPath,
            $"{Path.GetFileName(Path.TrimEndingDirectorySeparator(SourceFolderPath))}_合并结果.xlsx");
        Progress = 100;
        _notificationService.ShowSuccess("提示", "合并成功");
    }
    
    /// <summary>
    /// 合并多个工作簿，把多个工作表放到同一个工作簿中
    /// </summary>
    private async Task MultiSheetMergeAsync()
    {
        var filePaths = await _filePickerService.GetFolderFilesAsync(SourceFolderPath, [".xlsx"]);
        if (filePaths.Length == 0)
        {
            return;
        }
        Progress = 0;
        using var workbook = await Task.Run(() => _multiSheetMergeService.Merge(filePaths, _progressReporter));
        await _filePickerService.SaveExcelAsync(
            workbook,
            TargetFolderPath,
            $"{Path.GetFileName(Path.TrimEndingDirectorySeparator(SourceFolderPath))}_合并结果.xlsx");
        Progress = 100;
        _notificationService.ShowSuccess("提示", "合并成功");
    }

    /// <summary>
    /// 读取一个Excel中第一个sheet的第一列，根据这列创建文件夹
    /// </summary>
    private async Task CreateFolderFromExcelAsync()
    {
        var names = _createFolderFromExcelService.GetFolderNames(SourceFilePath);
        var parentFolder = Path.GetDirectoryName(SourceFilePath);
        if (parentFolder is null)
        {
            return;
        }
        await Task.Run(() => _createFolderFromExcelService.CreateFolders(TargetFolderPath, names, _progressReporter));
        _notificationService.ShowSuccess("提示", "创建成功");
    }
}
