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
    public async Task<string[]> PickExcelFilesAsync()
    {
        var files = await _storageProvider.OpenFilePickerAsync(
            new FilePickerOpenOptions
            {
                Title = "选择Excel文件",
                AllowMultiple = true,
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
    /// <returns></returns>
    public async Task<string[]> PickFolderFilesAsync(
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
            return [];
        }

        string folderPath = folders[0].Path.LocalPath;
        var searchOption = includeSubdirectories
            ? SearchOption.AllDirectories
            : SearchOption.TopDirectoryOnly;
        return Directory
            .EnumerateFiles(folderPath, "*", searchOption)
            .Where(file => !Path.GetFileName(file).StartsWith("~$"))
            .Where(file => extensions.Contains(
                Path.GetExtension(file),
                StringComparer.OrdinalIgnoreCase))
            .ToArray();
    }

    public async Task<string?> SaveExcelAsync(
        XLWorkbook workbook,
        string suggestedFileName = "合并结果.xlsx")
    {
        var file = await _storageProvider.SaveFilePickerAsync(
            new FilePickerSaveOptions
            {
                Title = "保存Excel文件",
                SuggestedFileName = suggestedFileName,
                DefaultExtension = "xlsx",
                ShowOverwritePrompt = true,
                FileTypeChoices = 
                    [
                        new FilePickerFileType("Excel 工作簿")
                        {
                            Patterns = ["*.xlsx"],
                            MimeTypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
                        }
                    ]
            });

        if (file is null)
        {
            return null;
        }

        await using var stream = await file.OpenWriteAsync();
        workbook.SaveAs(stream);

        return file.Path.LocalPath;
    }
}