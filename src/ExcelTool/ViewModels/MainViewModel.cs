using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    private readonly RdAttendanceGenerationService _rdAttendanceGenerationService;
    private readonly FilePickerService _filePickerService;
    private readonly INotificationService _notificationService;
    private readonly IProgress<double> _progressReporter;

    public IReadOnlyList<FunctionOption> Functions { get; } =
    [
        new(ExcelFunction.MergeWorkbooks, "合并工作簿"),
        new(ExcelFunction.MergeWorksheets, "合并工作表"),
        new(ExcelFunction.CreateFolders, "批量创建文件夹"),
        new(ExcelFunction.GenerateRdAttendance, "研发工时及日志"),
    ];

    [ObservableProperty] private double _progress;
    [ObservableProperty] private string _sourceFolderPath = "";
    [ObservableProperty] private string _targetFolderPath = "";
    [ObservableProperty] private FunctionOption _selectedFunction;
    [ObservableProperty] private string _sourceFilePath = "";
    [ObservableProperty] private string _rdAttendanceFilePath = "";
    [ObservableProperty] private string _rdProjectSummaryFilePath = "";
    [ObservableProperty] private string _rdPersonProjectFilePath = "";
    [ObservableProperty] private string _rdStaffFilePath = "";
    [ObservableProperty] private string _rdRulesFilePath = "";
    [ObservableProperty] private string _rdHistoryLockFilePath = "";
    [ObservableProperty] private string _rdControlTableFilePath = "";

    public string RdAttendanceDisplayPath => CompactPath(RdAttendanceFilePath);
    public string RdProjectSummaryDisplayPath => CompactPath(RdProjectSummaryFilePath);
    public string RdPersonProjectDisplayPath => CompactPath(RdPersonProjectFilePath);
    public string RdStaffDisplayPath => CompactPath(RdStaffFilePath);
    public string RdRulesDisplayPath => CompactPath(RdRulesFilePath);
    public string RdHistoryLockDisplayPath => CompactPath(RdHistoryLockFilePath);
    public string RdControlTableDisplayPath => CompactPath(RdControlTableFilePath);

    public bool ShowSelectSourceFolder => SelectedFunction.Value is ExcelFunction.MergeWorkbooks or ExcelFunction.MergeWorksheets;
    public bool ShowSelectTargetFolder => 
        SelectedFunction.Value is ExcelFunction.MergeWorkbooks or ExcelFunction.MergeWorksheets or ExcelFunction.CreateFolders or ExcelFunction.GenerateRdAttendance;
    public bool ShowSelectSourceFile => SelectedFunction.Value is ExcelFunction.CreateFolders;
    public bool ShowSelectSourceFiles => SelectedFunction.Value is ExcelFunction.GenerateRdAttendance;
    public string WorkspaceTitle => SelectedFunction.DisplayName;
    public string WorkspaceDescription => SelectedFunction.Value switch
    {
        ExcelFunction.MergeWorkbooks => "将文件夹内多个 Excel 工作簿的数据汇总到同一张工作表。",
        ExcelFunction.MergeWorksheets => "将多个 Excel 工作簿中的工作表合并到一个工作簿中。",
        ExcelFunction.CreateFolders => "读取 Excel 首个工作表的第一列，并按内容批量创建文件夹。",
        ExcelFunction.GenerateRdAttendance => "根据考勤、项目与人员资料生成研发工时及配套日志。",
        _ => "选择一项任务开始处理。"
    };
    public string ExecuteButtonText => SelectedFunction.Value switch
    {
        ExcelFunction.MergeWorkbooks => "开始合并",
        ExcelFunction.MergeWorksheets => "开始合并",
        ExcelFunction.CreateFolders => "开始创建",
        ExcelFunction.GenerateRdAttendance => "开始生成",
        _ => "开始执行"
    };
    public string ProgressStatus => Progress >= 100 ? "处理完成" : Progress > 0 ? "正在处理，请稍候…" : "就绪，等待开始任务。";

    partial void OnSelectedFunctionChanged(FunctionOption value)
    {
        OnPropertyChanged(nameof(ShowSelectSourceFolder));
        OnPropertyChanged(nameof(ShowSelectTargetFolder));
        OnPropertyChanged(nameof(ShowSelectSourceFile));
        OnPropertyChanged(nameof(ShowSelectSourceFiles));
        OnPropertyChanged(nameof(WorkspaceTitle));
        OnPropertyChanged(nameof(WorkspaceDescription));
        OnPropertyChanged(nameof(ExecuteButtonText));
    }

    partial void OnProgressChanged(double value) => OnPropertyChanged(nameof(ProgressStatus));
    
    
    public MainViewModel(
        MultiExcelMergeService multiExcelMergeService,
        MultiSheetMergeService multiSheetMergeService,
        CreateFolderFromExcelService createFolderFromExcelService,
        RdAttendanceGenerationService rdAttendanceGenerationService,
        FilePickerService filePickerService,
        INotificationService notificationService)
    {
        _multiExcelMergeService = multiExcelMergeService;
        _multiSheetMergeService = multiSheetMergeService;
        _createFolderFromExcelService = createFolderFromExcelService;
        _rdAttendanceGenerationService = rdAttendanceGenerationService;
        _filePickerService = filePickerService;
        _notificationService = notificationService;
        
        _selectedFunction = Functions[0];
        _progressReporter = new Progress<double>(value => Progress = value);
        TargetFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    }

    partial void OnRdAttendanceFilePathChanged(string value) => OnPropertyChanged(nameof(RdAttendanceDisplayPath));
    partial void OnRdProjectSummaryFilePathChanged(string value) => OnPropertyChanged(nameof(RdProjectSummaryDisplayPath));
    partial void OnRdPersonProjectFilePathChanged(string value) => OnPropertyChanged(nameof(RdPersonProjectDisplayPath));
    partial void OnRdStaffFilePathChanged(string value) => OnPropertyChanged(nameof(RdStaffDisplayPath));
    partial void OnRdRulesFilePathChanged(string value) => OnPropertyChanged(nameof(RdRulesDisplayPath));
    partial void OnRdHistoryLockFilePathChanged(string value) => OnPropertyChanged(nameof(RdHistoryLockDisplayPath));
    partial void OnRdControlTableFilePathChanged(string value) => OnPropertyChanged(nameof(RdControlTableDisplayPath));

    private static string CompactPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return path;

        var fileName = Path.GetFileName(path);
        return string.IsNullOrWhiteSpace(fileName) ? path : $"…/{fileName}";
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

    [RelayCommand]
    private async Task<string?> PickRdSourceFileAsync()
    {
        var filePaths = await _filePickerService.PickExcelFilesAsync(false);
        if (filePaths.Length == 0) return null;
        TargetFolderPath = Path.GetDirectoryName(filePaths[0]) ?? TargetFolderPath;
        return filePaths[0];
    }

    [RelayCommand] private async Task SelectRdAttendanceFileAsync() =>
        RdAttendanceFilePath = await PickRdSourceFileAsync() ?? RdAttendanceFilePath;

    [RelayCommand] private async Task SelectRdProjectSummaryFileAsync() =>
        RdProjectSummaryFilePath = await PickRdSourceFileAsync() ?? RdProjectSummaryFilePath;

    [RelayCommand] private async Task SelectRdPersonProjectFileAsync() =>
        RdPersonProjectFilePath = await PickRdSourceFileAsync() ?? RdPersonProjectFilePath;

    [RelayCommand] private async Task SelectRdStaffFileAsync() =>
        RdStaffFilePath = await PickRdSourceFileAsync() ?? RdStaffFilePath;

    [RelayCommand] private async Task SelectRdRulesFileAsync() =>
        RdRulesFilePath = await PickRdSourceFileAsync() ?? RdRulesFilePath;

    [RelayCommand] private async Task SelectRdHistoryLockFileAsync() =>
        RdHistoryLockFilePath = await PickRdSourceFileAsync() ?? RdHistoryLockFilePath;

    [RelayCommand] private async Task SelectRdControlTableFileAsync() =>
        RdControlTableFilePath = await PickRdSourceFileAsync() ?? RdControlTableFilePath;

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


        if (ShowSelectSourceFiles && new[]
            {
                RdAttendanceFilePath, RdProjectSummaryFilePath, RdPersonProjectFilePath,
                RdStaffFilePath, RdRulesFilePath
            }.Any(string.IsNullOrWhiteSpace))
        {
            _notificationService.ShowWarning("提示", "请分别选择全部 5 类必需 Excel 文件");
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
            case ExcelFunction.GenerateRdAttendance:
                await GenerateRdAttendanceAsync();
                break;
            default:
                _notificationService.ShowWarning("提示", "请选择需要执行的功能");
                break;
        }
    }

    private async Task GenerateRdAttendanceAsync()
    {
        Progress = 0;
        try
        {
            var result = await Task.Run(() => _rdAttendanceGenerationService.Generate(
                new RdAttendanceSourceFiles(
                    RdAttendanceFilePath,
                    RdProjectSummaryFilePath,
                    RdPersonProjectFilePath,
                    RdStaffFilePath,
                    RdRulesFilePath,
                    RdHistoryLockFilePath,
                    RdControlTableFilePath),
                TargetFolderPath,
                _progressReporter));
            Progress = 100;
            _notificationService.ShowSuccess(
                "生成成功",
                $"已生成内部复核版、发企业版及 {result.LogFileCount} 份研发日志\n保存位置：{result.OutputFolder}");
        }
        catch (Exception ex)
        {
            _notificationService.ShowWarning("生成失败", ex.Message);
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
