using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelTool.Services.Workflows;

public class MultiExcelMergeService
{
    public XLWorkbook Merge(IEnumerable<string> filePaths)
    {
        var targetWorkbook = new XLWorkbook();
        int index = 0;
        // 遍历所有Excel
        foreach (var filePath in filePaths)
        {
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

                    if (firstColumn is null || lastColumn is null || (index > 0 && sourceRow.RowNumber() == 1))
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
            index++;
        }

        return targetWorkbook;
    }
}