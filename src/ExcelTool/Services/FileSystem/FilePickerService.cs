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
    /// 选择文件夹并返回路径
    /// </summary>
    /// <param name="allowMultiple"></param>
    /// <returns></returns>
    public async Task<string[]> PickFolderAsync(bool allowMultiple = false)
    {
        var folders = await _storageProvider.OpenFolderPickerAsync(
            new FolderPickerOpenOptions
            {
                Title = "选择文件夹",
                AllowMultiple = allowMultiple,
            });
        return folders.Select(folder => folder.Path.LocalPath).ToArray();
    }

    /// <summary>
    /// 选择指定格式的文件并返回路径
    /// </summary>
    /// <param name="extensions"></param>
    /// <param name="allowMultiple"></param>
    /// <returns></returns>
    public async Task<string[]> PickFileAsync(
        string[] extensions,
        bool allowMultiple = false)
    {
        var files = await _storageProvider.OpenFilePickerAsync(
            new FilePickerOpenOptions
            {
                Title = "选择文件",
                AllowMultiple = allowMultiple,
                FileTypeFilter = 
                    [
                        new FilePickerFileType("自定义文件")
                        {
                            Patterns = extensions
                        }
                    ]
            });
        return files.Select(file => file.Path.LocalPath).ToArray();
    }

    /// <summary>
    /// 获取文件夹下所有指定类型的文件并返回路径
    /// </summary>
    /// <param name="folderPath"></param>
    /// <param name="extensions"></param>
    /// <param name="allowMultiple"></param>
    /// <returns></returns>
    public async Task<string[]> GetFolderFilesAsync(
        string folderPath,
        string[] extensions,
        bool allowMultiple = false)
    {
        var searchOption = allowMultiple ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        return Directory
            .EnumerateFiles(folderPath, "*", searchOption)
            .Where(file => !Path.GetFileName(file).StartsWith("~$"))
            .Where(file => extensions.Contains(
                Path.GetExtension(file),
                StringComparer.OrdinalIgnoreCase))
            .ToArray();
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
    /// 保存Excel到指定位置
    /// </summary>
    /// <param name="workbook"></param>
    /// <param name="targetFolderPath"></param>
    /// <param name="suggestedFileName"></param>
    /// <returns></returns>
    public async Task<string?> SaveExcelAsync(
        XLWorkbook workbook,
        string targetFolderPath,
        string suggestedFileName = "合并结果.xlsx")
    {
        var filePath = Path.Combine(targetFolderPath, suggestedFileName);
        workbook.SaveAs(filePath);

        return filePath;
    }
}
