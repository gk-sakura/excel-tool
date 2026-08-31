using System;
using System.IO;
using System.Linq;
using Avalonia.Platform.Storage;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelTool.Services.FileSystem;

public class FilePickerService
{
    private readonly IStorageProvider _storageProvider;

    public FilePickerService(IStorageProvider storageProvider)
    {
        _storageProvider = storageProvider;
    }

    /// <summary>
    /// 选择多个Excel文件
    /// </summary>
    /// <returns></returns>
    public async Task<string[]> PickExcelFilesAsync(bool allowMultiple = true)
    {
        var files = await _storageProvider.OpenFilePickerAsync(
            new FilePickerOpenOptions
            {
                Title = "选择Excel文件",
                AllowMultiple = allowMultiple,
                FileTypeFilter = [
                    new FilePickerFileType("Excel 文件")
                    {
                        Patterns = ["*.xlsx"]
                    }
                ]
            });
        
        return files.Select(file => file.Path.LocalPath).ToArray();
    }

    /// <summary>
    /// 选择文件夹，并读取指定格式的文件
    /// </summary>
    /// <param name="extensions"></param>
    /// <param name="includeSubdirectories"></param>
    /// <returns></returns>
    public async Task<(string[] FilePaths, string? FolderPath)> PickFolderFilesAsync(
        string[] extensions,
        bool includeSubdirectories = false)
    {
        var folders = await _storageProvider.OpenFolderPickerAsync(
            new FolderPickerOpenOptions
            {
                Title = "选择文件夹",
                AllowMultiple = false
            });
        // 取消选择
        if (folders.Count == 0)
        {
            return ([], null);
        }

        string folderPath = folders[0].Path.LocalPath;
        var searchOption = includeSubdirectories
            ? SearchOption.AllDirectories
            : SearchOption.TopDirectoryOnly;
        var filePaths = Directory
            .EnumerateFiles(folderPath, "*", searchOption)
            .Where(file => !Path.GetFileName(file).StartsWith("~$"))
            .Where(file => extensions.Contains(
                Path.GetExtension(file),
                StringComparer.OrdinalIgnoreCase))
            .ToArray();

        return (filePaths, folderPath);
    }

    /// <summary>
    /// 保存Excel到桌面
    /// </summary>
    /// <param name="workbook"></param>
    /// <param name="suggestedFileName"></param>
    /// <returns></returns>
    public async Task<string?> SaveExcelAsync(
        XLWorkbook workbook,
        string suggestedFileName = "合并结果.xlsx")
    {
        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var filePath = Path.Combine(desktopPath, suggestedFileName);
        workbook.SaveAs(filePath);

        return filePath;
    }
}
