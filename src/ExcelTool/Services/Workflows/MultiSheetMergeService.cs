using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class MultiSheetMergeService
{
    public void Merge(IEnumerable<string> filePaths)
    {
        using var workbook = new XLWorkbook();
    }
}