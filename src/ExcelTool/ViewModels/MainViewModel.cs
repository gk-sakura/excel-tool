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
    private readonly FilePickerService _filePickerService;
    
    public MainViewModel(
        MultiExcelMergeService multiExcelMergeService,
        MultiSheetMergeService multiSheetMergeService,
        FilePickerService filePickerService)
    {
        _multiExcelMergeService = multiExcelMergeService;
        _multiSheetMergeService = multiSheetMergeService;
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
}
