
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
}