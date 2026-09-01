using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class MultiExcelMergeService
{
    public XLWorkbook Merge(
        IEnumerable<string> filePaths,
        IProgress<double>? progress = null)
    {
        var targetWorkbook = new XLWorkbook();
        var fileList = filePaths.ToArray();
        if (fileList.Length == 0)
        {
            return targetWorkbook;
        }
        // 遍历所有Excel
        for(var index = 0; index < fileList.Length; index++)
        {
            var filePath = fileList[index];
            using var sourceWorkbook = new XLWorkbook(filePath);
            // 遍历当前Excel的所有sheet
            foreach (var sourceSheet in sourceWorkbook.Worksheets)
            {
                var isNewTargetSheet = !targetWorkbook.Worksheets.Contains(sourceSheet.Name);
                if (isNewTargetSheet)
                {
                    targetWorkbook.Worksheets.Add(sourceSheet.Name);
                }
                var targetSheet = targetWorkbook.Worksheet(sourceSheet.Name);
                // 复制列宽
                var lastColumnNumber = sourceSheet.LastColumnUsed()?.ColumnNumber() ?? 0;
                for (var columnNumber = 1; columnNumber <= lastColumnNumber; columnNumber++)
                {
                    var sourceWidth = sourceSheet.Column(columnNumber).Width;
                    var targetColumn = targetSheet.Column(columnNumber);
                    targetColumn.Width = isNewTargetSheet ? sourceWidth : Math.Max(targetColumn.Width, sourceWidth);
                }
                // 遍历当前sheet的所有行
                var targetRowNumber = (targetSheet.LastRowUsed()?.RowNumber() ?? 0) + 1;
                Console.WriteLine(targetRowNumber);
                foreach (var sourceRow in sourceSheet.Rows())
                {
                    var firstColumn = sourceRow.FirstCellUsed()?.Address.ColumnNumber;
                    var lastColumn = sourceRow.LastCellUsed()?.Address.ColumnNumber;

                    if (firstColumn is null || lastColumn is null)
                    {
                        continue;
                    }
                    var sourceRange = sourceSheet.Range(
                        sourceRow.RowNumber(),
                        firstColumn.Value,
                        sourceRow.RowNumber(),
                        lastColumn.Value);
                    sourceRange.CopyTo(targetSheet.Cell(targetRowNumber, firstColumn.Value));
                    targetSheet.Row(targetRowNumber).Height = sourceRow.Height;
                    targetRowNumber++;
                }
            }
            
            progress?.Report((index + 1) * 100.0 / fileList.Length);
        }

        return targetWorkbook;
    }
}