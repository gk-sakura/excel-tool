using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class MultiSheetMergeService
{
    public XLWorkbook Merge(IEnumerable<string> filePaths)
    {
        var targetWorkbook = new XLWorkbook();
        foreach (var filepath in filePaths)
        {
            var sourceWorkbook = new XLWorkbook(filepath);
            foreach (var sourceSheet in sourceWorkbook.Worksheets)
            {
                sourceSheet.CopyTo(targetWorkbook, $"{Path.GetFileNameWithoutExtension(filepath)}{sourceSheet.Name}");
            }
        }

        return targetWorkbook;
    }
}