using System;
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
        var filePaths = await _filePickerService.PickFolderFilesAsync([".xlsx"]);
        if (filePaths.Length == 0)
        {
            return;
        }
        using var workbook = _multiExcelMergeService.Merge(filePaths);
        var savedPath = await _filePickerService.SaveExcelAsync(
            workbook,
            $"合并结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        if (savedPath is null)
        {
            return;
        }
    }
    
    // 合并同一个Excel中的多个sheet，可以选择多个Excel文件分别合并
    [RelayCommand]
    private async Task MultiSheetMergeAsync()
    {
        var filePaths = await _filePickerService.PickFolderFilesAsync([".xlsx"]);
        if (filePaths.Length == 0)
        {
            return;
        }
        _multiSheetMergeService.Merge(filePaths);
    }
}
