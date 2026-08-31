using System;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelTool.Services.FileSystem;
using ExcelTool.Services.Workflows;
using System.Threading.Tasks;

namespace ExcelTool.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private readonly MultiExcelMergeService _multiExcelMergeService;
    private readonly MultiSheetMergeService _multiSheetMergeService;
    private readonly CreateFolderFromExcelService _createFolderFromExcelService;
    private readonly FilePickerService _filePickerService;
    
    public MainViewModel(
        MultiExcelMergeService multiExcelMergeService,
        MultiSheetMergeService multiSheetMergeService,
        CreateFolderFromExcelService createFolderFromExcelService,
        FilePickerService filePickerService)
    {
        _multiExcelMergeService = multiExcelMergeService;
        _multiSheetMergeService = multiSheetMergeService;
        _createFolderFromExcelService = createFolderFromExcelService;
        _filePickerService = filePickerService;
    }
    
    // 合并多个Excel
    [RelayCommand]
    private async Task MultiExcelMergeAsync()
    {
        var (filePaths, folderPath) = await _filePickerService.PickFolderFilesAsync([".xlsx"]);
        Console.WriteLine(folderPath);
        if (filePaths.Length == 0 || folderPath is null)
        {
            return;
        }
        using var workbook = _multiExcelMergeService.Merge(filePaths);
        await _filePickerService.SaveExcelAsync(
            workbook,
            $"{Path.GetFileName(Path.TrimEndingDirectorySeparator(folderPath))}_合并结果.xlsx");
    }
    
    // 合并同一个Excel中的多个sheet，可以选择多个Excel文件分别合并
    [RelayCommand]
    private async Task MultiSheetMergeAsync()
    {
        var (filePaths, folderPath) = await _filePickerService.PickFolderFilesAsync([".xlsx"]);
        if (filePaths.Length == 0 || folderPath is null)
        {
            return;
        }
        using var workbook = _multiSheetMergeService.Merge(filePaths);
        await _filePickerService.SaveExcelAsync(
            workbook,
            $"{Path.GetFileName(Path.TrimEndingDirectorySeparator(folderPath))}_合并结果.xlsx");
    }

    /// <summary>
    /// 读取一个Excel中第一个sheet的第一列，根据这列创建文件夹
    /// </summary>
    [RelayCommand]
    private async Task CreateFolderFromExcelAsync()
    {
        var filePaths = await _filePickerService.PickExcelFilesAsync(false);
        if (filePaths.Length == 0)
        {
            return;
        }
        var names = _createFolderFromExcelService.GetFolderNames(filePaths[0]);
        var parentFolder = Path.GetDirectoryName(filePaths[0]);
        if (parentFolder is null)
        {
            return;
        }
        foreach (var name in names)
        {
            Directory.CreateDirectory(Path.Combine(parentFolder, name));
        }
    }
}
