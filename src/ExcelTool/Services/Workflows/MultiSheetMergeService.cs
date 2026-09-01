using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class MultiSheetMergeService
{
    public XLWorkbook Merge(
        IEnumerable<string> filePaths,
        IProgress<double>? progress = null)
    {
        var targetWorkbook = new XLWorkbook();
        var fileList = filePaths.ToArray();
        for (var index = 0; index < fileList.Length; index++)
        {
            var filePath = fileList[index];
            var sourceWorkbook = new XLWorkbook(filePath);
            foreach (var sourceSheet in sourceWorkbook.Worksheets)
            {
                sourceSheet.CopyTo(targetWorkbook, $"{Path.GetFileNameWithoutExtension(filePath)}{sourceSheet.Name}");
            }
            progress?.Report((index + 1) * 100.0 / fileList.Length);
        }

        return targetWorkbook;
    }
}