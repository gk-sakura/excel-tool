
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class CreateFolderFromExcelService
{
    public string[] GetFolderNames(string filePath)
    {
        var sourceWorkbook = new XLWorkbook(filePath);
        var sourceSheet = sourceWorkbook.Worksheets.First();
        return sourceSheet
            .Column(1)
            .CellsUsed()
            .Select(cell => cell.GetString())
            .ToArray();
    }

    public void CreateFolders(
        string targetFolderPath,
        IEnumerable<string> folderNames,
        IProgress<double>? progress = null)
    {
        var folders = folderNames.ToArray();
        for (var i = 0; i < folders.Length; i++)
        {
            Directory.CreateDirectory(Path.Combine(targetFolderPath, folders[i]));
            progress?.Report((i + 1) * 100.0 / folders.Length);
        }
    }
}