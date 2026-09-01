using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelDataReader;

namespace ExcelTool.Services.Workflows;

public sealed record RdAttendanceGenerationResult(string OutputFolder, int LogFileCount);
public sealed record RdAttendanceSourceFiles(
    string Attendance,
    string ProjectSummary,
    string PersonProject,
    string Staff,
    string Rules,
    string? HistoryLock = null,
    string? ControlTable = null);

public sealed class RdAttendanceGenerationService
{
    private static readonly string[] MatrixSheets =
    [
        "全职人员考勤源表", "全职人员考勤处理表", "全职人员分项目情况表", "全职人员考勤工时表"
    ];

    private sealed record Person(string Seq, string EmployeeId, string Name, string Position);
    private sealed record AttendanceRule(string Normalized, double Days, bool IsValid);
    private sealed record Project(string Code, string Name, DateTime? Start, DateTime? End);
    private sealed record InputFiles(string Attendance, string ProjectSummary, string PersonProject, string Staff, string Rules,
        string? HistoryLock = null, string? ControlTable = null);
    private sealed record AllocationOptions(bool Enabled, int MinimumDays, int MaximumDays, bool InvalidBreaksBlock,
        bool ContinueAcrossMonths, int? MaximumSameProjectDays, bool AvoidFrequentSwitching);
    private sealed record ProjectPeriodLock(string EmployeeId, string Name, string ProjectCode, DateTime Start, DateTime End, bool Force);
    private sealed record LogOptions(bool Daily, bool Weekly, bool Monthly);

    public RdAttendanceGenerationResult Generate(
        RdAttendanceSourceFiles sourceFiles,
        string targetFolder,
        IProgress<double>? progress = null)
    {
        var originals = new[]
        {
            sourceFiles.Attendance, sourceFiles.ProjectSummary, sourceFiles.PersonProject,
            sourceFiles.Staff, sourceFiles.Rules, sourceFiles.HistoryLock ?? "", sourceFiles.ControlTable ?? ""
        };
        var (normalizedFiles, temporaryFolder) = NormalizeSourceFiles(originals);
        try
        {
            var typedFiles = new InputFiles(
                normalizedFiles[0], normalizedFiles[1], normalizedFiles[2], normalizedFiles[3], normalizedFiles[4],
                NullIfEmpty(normalizedFiles[5]), NullIfEmpty(normalizedFiles[6]));
            return GenerateCore(normalizedFiles, targetFolder, progress, typedFiles);
        }
        finally
        {
            if (temporaryFolder is not null && Directory.Exists(temporaryFolder))
                Directory.Delete(temporaryFolder, true);
        }
    }

    public RdAttendanceGenerationResult Generate(
        IEnumerable<string> sourceFiles,
        string targetFolder,
        IProgress<double>? progress = null)
    {
        var originalFiles = sourceFiles.ToArray();
        var (normalizedFiles, temporaryFolder) = NormalizeSourceFiles(originalFiles);
        try
        {
            return GenerateCore(normalizedFiles, targetFolder, progress);
        }
        finally
        {
            if (temporaryFolder is not null && Directory.Exists(temporaryFolder))
                Directory.Delete(temporaryFolder, true);
        }
    }

    private RdAttendanceGenerationResult GenerateCore(
        IEnumerable<string> sourceFiles,
        string targetFolder,
        IProgress<double>? progress = null,
        InputFiles? selectedFiles = null)
    {
        var files = selectedFiles ?? IdentifyFiles(sourceFiles);
        progress?.Report(5);

        var people = ReadPeople(files.Staff);
        if (people.Count == 0)
            throw new InvalidDataException("花名册中未找到“是否纳入研发工时归集”为“是/全职/专职”的在职人员。");

        var (dates, attendance, isPunchLog) = ReadAttendance(files.Attendance, people.Select(x => x.Name).ToHashSet());
        if (dates.Count == 0)
            throw new InvalidDataException("考勤源表中未识别到日期列，请确认日期与“姓名”位于同一标题行。");

        var rules = ReadRules(files.Rules);
        if (isPunchLog && !rules.ContainsKey("上班"))
            rules["上班"] = new AttendanceRule("正常出勤", 1, true);
        var missingStatuses = attendance.Values.SelectMany(x => x.Values)
            .Where(x => !string.IsNullOrWhiteSpace(x) && !rules.ContainsKey(x))
            .Distinct(StringComparer.Ordinal).Order().ToArray();
        if (missingStatuses.Length > 0)
            throw new InvalidDataException("考勤规则中缺少或未确认以下状态：" + string.Join("、", missingStatuses));

        var projects = ReadProjects(files.ProjectSummary);
        var relations = ReadPersonProjectMap(files.PersonProject);
        var lockedMonths = ReadLockedMonths(files.HistoryLock);
        var (logOptions, allocationOptions, projectLocks) = ReadControlTable(files.ControlTable);
        var assignments = AssignProjects(dates, people, attendance, rules, projects, relations,
            lockedMonths, allocationOptions, projectLocks);
        progress?.Report(30);

        var metadata = DetectMetadata(sourceFiles, dates);
        var outputRoot = Path.Combine(targetFolder, $"研发工时及日志_{metadata.Company}_{metadata.Period}_{metadata.GeneratedOn}");
        Directory.CreateDirectory(outputRoot);

        var internalPath = Path.Combine(outputRoot,
            $"1.01 {metadata.Company} {metadata.Period}研发人员考勤情况及工时表（{metadata.GeneratedOn}内部复核）.xlsx");
        var enterprisePath = Path.Combine(outputRoot,
            $"2.01 {metadata.Company} {metadata.Period}研发人员考勤情况及工时表（{metadata.GeneratedOn}发企业）.xlsx");

        using (var workbook = LoadTemplate("InternalReview.xlsx"))
        {
            FillMatrixSheets(workbook, dates, people, attendance, rules, assignments);
            FillInternalReviewSheets(workbook, files.Rules, dates, people, attendance, rules, assignments, projects, relations, lockedMonths);
            workbook.SaveAs(internalPath);
        }
        progress?.Report(50);

        using (var workbook = LoadTemplate("Enterprise.xlsx"))
        {
            FillMatrixSheets(workbook, dates, people, attendance, rules, assignments);
            workbook.SaveAs(enterprisePath);
        }
        progress?.Report(65);

        var logCount = GenerateLogs(outputRoot, metadata.Period, dates, people, assignments, projects, logOptions, progress);
        return new RdAttendanceGenerationResult(outputRoot, logCount);
    }

    private static (string[] Files, string? TemporaryFolder) NormalizeSourceFiles(string[] sourceFiles)
    {
        var legacyFiles = sourceFiles
            .Where(path => Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (legacyFiles.Length == 0)
            return (sourceFiles, null);

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var temporaryFolder = Path.Combine(Path.GetTempPath(), $"ExcelTool-xls-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryFolder);
        var normalized = new List<string>(sourceFiles.Length);
        try
        {
            foreach (var sourceFile in sourceFiles)
            {
                if (!Path.GetExtension(sourceFile).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    normalized.Add(sourceFile);
                    continue;
                }

                var outputName = Path.GetFileNameWithoutExtension(sourceFile) + ".xlsx";
                var outputPath = Path.Combine(temporaryFolder, outputName);
                var suffix = 1;
                while (File.Exists(outputPath))
                    outputPath = Path.Combine(temporaryFolder, $"{Path.GetFileNameWithoutExtension(sourceFile)}-{suffix++}.xlsx");
                ConvertLegacyWorkbook(sourceFile, outputPath);
                normalized.Add(outputPath);
            }
            return (normalized.ToArray(), temporaryFolder);
        }
        catch
        {
            Directory.Delete(temporaryFolder, true);
            throw;
        }
    }

    private static HashSet<string> ReadLockedMonths(string? path)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return result;
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var header = FindHeaderRow(ws, "期间", "是否锁定");
        var map = HeaderMap(header);
        foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
        {
            if (Value(row, map, "是否锁定") != "是") continue;
            var period = NormalizeMonth(Value(row, map, "期间"));
            if (period.Length > 0) result.Add(period);
        }
        return result;
    }

    private static (LogOptions Logs, AllocationOptions Allocation, List<ProjectPeriodLock> Locks) ReadControlTable(string? path)
    {
        var logs = new LogOptions(true, true, true);
        var allocation = new AllocationOptions(true, 3, 4, false, true, 25, true);
        var locks = new List<ProjectPeriodLock>();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return (logs, allocation, locks);
        using var wb = new XLWorkbook(path);

        if (wb.TryGetWorksheet("研发日志生成情况", out var logSheet))
        {
            var header = FindHeaderRow(logSheet, "日志类型", "是否生成");
            var map = HeaderMap(header);
            var values = new Dictionary<string, bool>(StringComparer.Ordinal);
            foreach (var row in logSheet.Rows(header.RowNumber() + 1, logSheet.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
            {
                var type = Value(row, map, "日志类型");
                if (type.Length > 0) values[type] = Value(row, map, "是否生成") != "否";
            }
            logs = new LogOptions(values.GetValueOrDefault("日报", true), values.GetValueOrDefault("周报", true),
                values.GetValueOrDefault("月报", true));
        }

        if (wb.TryGetWorksheet("全局分配规则", out var allocationSheet))
        {
            var header = FindHeaderRow(allocationSheet, "参数", "默认值");
            var map = HeaderMap(header);
            var values = allocationSheet.Rows(header.RowNumber() + 1, allocationSheet.LastRowUsed()?.RowNumber() ?? header.RowNumber())
                .Where(row => Value(row, map, "参数").Length > 0)
                .ToDictionary(row => Value(row, map, "参数"), row => Value(row, map, "默认值"), StringComparer.Ordinal);
            allocation = new AllocationOptions(
                values.GetValueOrDefault("是否启用连续分配", "是") != "否",
                PositiveInt(values.GetValueOrDefault("默认最小连续有效工作日"), 3),
                PositiveInt(values.GetValueOrDefault("默认最大连续有效工作日"), 4),
                values.GetValueOrDefault("无效考勤是否打断连续块", "否") == "是",
                values.GetValueOrDefault("跨月是否延续连续块", "是") != "否",
                NullablePositiveInt(values.GetValueOrDefault("同项目最大连续有效工作日")),
                values.GetValueOrDefault("是否避免频繁来回切换", "是") != "否");
            if (allocation.MaximumDays < allocation.MinimumDays)
                allocation = allocation with { MaximumDays = allocation.MinimumDays };
        }

        if (wb.TryGetWorksheet("项目期间锁定", out var lockSheet))
        {
            var header = FindHeaderRow(lockSheet, "研发项目", "起始日期", "结束日期", "是否启用");
            var map = HeaderMap(header);
            foreach (var row in lockSheet.Rows(header.RowNumber() + 1, lockSheet.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
            {
                if (Value(row, map, "是否启用") != "是") continue;
                var project = Value(row, map, "研发项目");
                var start = ReadDate(row.Cell(map["起始日期"]));
                var end = ReadDate(row.Cell(map["结束日期"]));
                if (project.Length == 0 || !start.HasValue || !end.HasValue) continue;
                locks.Add(new ProjectPeriodLock(Value(row, map, "工号"), Value(row, map, "姓名"), project,
                    start.Value.Date, end.Value.Date, Value(row, map, "是否强制") == "是"));
            }
        }
        return (logs, allocation, locks);
    }

    private static void ConvertLegacyWorkbook(string sourcePath, string outputPath)
    {
        try
        {
            using var stream = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            });
            using var workbook = new XLWorkbook();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            do
            {
                var baseName = SanitizeWorksheetName(reader.Name);
                var sheetName = UniqueWorksheetName(baseName, usedNames);
                var worksheet = workbook.Worksheets.Add(sheetName);
                var rowNumber = 1;
                while (reader.Read())
                {
                    for (var column = 0; column < reader.FieldCount; column++)
                    {
                        var value = reader.GetValue(column);
                        if (value is null or DBNull) continue;
                        var cell = worksheet.Cell(rowNumber, column + 1);
                        cell.Value = XLCellValue.FromObject(value);
                        var numberFormat = reader.GetNumberFormatString(column);
                        if (!string.IsNullOrWhiteSpace(numberFormat))
                            cell.Style.NumberFormat.Format = numberFormat;
                    }
                    rowNumber++;
                }
            } while (reader.NextResult());

            if (!workbook.Worksheets.Any())
                throw new InvalidDataException("工作簿中没有可读取的工作表");
            workbook.SaveAs(outputPath);
        }
        catch (Exception ex)
        {
            throw new InvalidDataException($"旧版 Excel 文件“{Path.GetFileName(sourcePath)}”转换失败：{ex.Message}", ex);
        }
    }

    private static string SanitizeWorksheetName(string? name)
    {
        var result = Regex.Replace(string.IsNullOrWhiteSpace(name) ? "Sheet" : name, @"[\\/:?*\[\]]", "_");
        return result.Length <= 31 ? result : result[..31];
    }

    private static string UniqueWorksheetName(string baseName, HashSet<string> usedNames)
    {
        var name = baseName;
        var suffix = 1;
        while (!usedNames.Add(name))
        {
            var tail = $"-{suffix++}";
            name = baseName[..Math.Min(baseName.Length, 31 - tail.Length)] + tail;
        }
        return name;
    }

    private static InputFiles IdentifyFiles(IEnumerable<string> paths)
    {
        string? attendance = null, project = null, relation = null, staff = null, rules = null;
        foreach (var path in paths.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!File.Exists(path) || !Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                continue;
            try
            {
                using var wb = new XLWorkbook(path);
                foreach (var ws in wb.Worksheets)
                {
                    foreach (var row in ws.RowsUsed().Take(20))
                    {
                        var headers = row.Cells(1, Math.Max(1, ws.LastColumnUsed()?.ColumnNumber() ?? 1))
                            .Select(CellText).ToArray();
                        var set = headers.ToHashSet(StringComparer.Ordinal);
                        if (set.Contains("项目编码") && set.Contains("项目名称")) project ??= path;
                        if (set.Contains("原始考勤状态") && set.Contains("折算天数")) rules ??= path;
                        if (set.Contains("是否纳入研发工时归集") && set.Contains("姓名")) staff ??= path;
                        if (set.Contains("姓名") && headers.Any(x => Regex.IsMatch(x, "^RD\\d+", RegexOptions.IgnoreCase))) relation ??= path;
                        if (set.Any(x => x is "姓名" or "员工姓名" or "人员姓名") && row.CellsUsed().Count(IsDateCell) >= 2) attendance ??= path;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"无法读取源文件“{Path.GetFileName(path)}”：{ex.Message}", ex);
            }
        }

        var missing = new List<string>();
        if (attendance is null) missing.Add("研发人员考勤源表");
        if (project is null) missing.Add("研发项目及周期汇总表");
        if (relation is null) missing.Add("研发人员与研发项目关系匹配表");
        if (staff is null) missing.Add("研发人员花名册及基础信息");
        if (rules is null) missing.Add("考勤规则拆解表");
        if (missing.Count > 0)
            throw new InvalidDataException("未能从所选文件中识别：" + string.Join("、", missing) + "。请一次性选择这 5 类 Excel 文件。");
        return new InputFiles(attendance!, project!, relation!, staff!, rules!);
    }

    private static List<Person> ReadPeople(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var headerRow = FindHeaderRow(ws, "姓名");
        var headers = HeaderMap(headerRow);
        var result = new List<Person>();
        foreach (var row in ws.Rows(headerRow.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? headerRow.RowNumber()))
        {
            var name = CellText(row.Cell(headers["姓名"]));
            if (name.Length == 0) continue;
            var include = headers.TryGetValue("是否纳入研发工时归集", out var includeCol)
                ? CellText(row.Cell(includeCol)) : "是";
            var status = headers.TryGetValue("状态", out var statusCol) ? CellText(row.Cell(statusCol)) : "";
            if (include is not ("是" or "全职" or "专职" or "专职研发") || status == "离职") continue;
            var employeeId = Value(row, headers, "工号");
            if (result.Any(x => (employeeId.Length > 0 && x.EmployeeId == employeeId) || x.Name == name)) continue;
            result.Add(new Person(
                Value(row, headers, "序号").Length > 0 ? Value(row, headers, "序号") : (result.Count + 1).ToString(),
                employeeId, name, Value(row, headers, "岗位").Length > 0 ? Value(row, headers, "岗位") : Value(row, headers, "岗位名称")));
        }
        return result;
    }

    private static (List<DateTime> Dates, Dictionary<string, Dictionary<DateTime, string>> Data, bool IsPunchLog) ReadAttendance(
        string path, HashSet<string> people)
    {
        using var wb = new XLWorkbook(path);
        var result = people.ToDictionary(x => x, _ => new Dictionary<DateTime, string>(), StringComparer.Ordinal);
        var allDates = new SortedSet<DateTime>();
        var foundMatrix = false;
        foreach (var ws in wb.Worksheets)
        {
            IXLRow? bestRow = null;
            List<(int Col, DateTime Date)> bestDates = [];
            foreach (var row in ws.RowsUsed())
            {
                var found = row.Cells(1, Math.Max(1, ws.LastColumnUsed()?.ColumnNumber() ?? 1))
                    .Where(IsDateCell).Select(c => (c.Address.ColumnNumber, Date: ReadDate(c)!.Value.Date)).ToList();
                if (found.Count > bestDates.Count) { bestRow = row; bestDates = found; }
            }
            if (bestRow is null || bestDates.Count == 0) continue;
            var nameCell = bestRow.CellsUsed().FirstOrDefault(c => CellText(c) is "姓名" or "员工姓名" or "人员姓名" or "名称");
            if (nameCell is null) continue;
            foundMatrix = true;
            var endRow = ws.LastRowUsed()?.RowNumber() ?? bestRow.RowNumber();
            for (var r = bestRow.RowNumber() + 1; r <= endRow; r++)
            {
                var name = CellText(ws.Cell(r, nameCell.Address.ColumnNumber));
                if (!people.Contains(name)) continue;
                foreach (var (col, date) in bestDates)
                {
                    var status = CellText(ws.Cell(r, col));
                    if (result[name].TryGetValue(date, out var old) && old != status)
                        throw new InvalidDataException($"同一人员同一日期存在矛盾考勤记录：{name} {date:yyyy-MM-dd}");
                    result[name][date] = status;
                    allDates.Add(date);
                }
            }
        }
        if (foundMatrix) return (allDates.ToList(), result, false);

        foreach (var ws in wb.Worksheets)
        {
            var header = ws.RowsUsed().FirstOrDefault(row =>
                row.CellsUsed().Any(c => CellText(c) is "姓名" or "员工姓名") &&
                row.CellsUsed().Any(c => CellText(c) is "打卡时间" or "签到时间"));
            if (header is null) continue;
            var map = HeaderMap(header);
            var nameColumn = map.GetValueOrDefault("姓名", map.GetValueOrDefault("员工姓名"));
            var timeColumn = map.GetValueOrDefault("打卡时间", map.GetValueOrDefault("签到时间"));
            foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
            {
                var name = CellText(row.Cell(nameColumn));
                var date = ReadDate(row.Cell(timeColumn))?.Date;
                if (!people.Contains(name) || !date.HasValue) continue;
                result[name][date.Value] = "上班";
                allDates.Add(date.Value);
            }
        }
        return (allDates.ToList(), result, true);
    }

    private static Dictionary<string, AttendanceRule> ReadRules(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.FirstOrDefault(x => x.Name == "考勤规则拆解表") ?? wb.Worksheets.First();
        var header = FindHeaderRow(ws, "原始考勤状态", "折算天数");
        var map = HeaderMap(header);
        var result = new Dictionary<string, AttendanceRule>(StringComparer.Ordinal);
        foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
        {
            var raw = Value(row, map, "原始考勤状态");
            if (raw.Length == 0 || Value(row, map, "是否已确认") != "是") continue;
            var normalized = Value(row, map, "规范考勤状态");
            if (!double.TryParse(Value(row, map, "折算天数"), NumberStyles.Any, CultureInfo.InvariantCulture, out var days)) days = 0;
            result[raw] = new AttendanceRule(normalized.Length == 0 ? raw : normalized, days, Value(row, map, "有效规则") == "有效");
        }
        return result;
    }

    private static Dictionary<string, Project> ReadProjects(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var header = ws.RowsUsed().FirstOrDefault(row =>
            row.CellsUsed().Any(c => CellText(c) == "项目名称") &&
            row.CellsUsed().Any(c => CellText(c) is "项目编码" or "项目编号-外部"))
            ?? throw new InvalidDataException($"工作表“{ws.Name}”缺少项目编号和项目名称列");
        var map = HeaderMap(header);
        var result = new Dictionary<string, Project>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
        {
            var code = Value(row, map, "项目编码");
            if (code.Length == 0) code = Value(row, map, "项目编号-外部");
            if (code.Length == 0) continue;
            var period = Value(row, map, "项目周期");
            if (period.Length == 0) period = Value(row, map, "项目起止时间");
            var (start, end) = ParsePeriod(period);
            result[code] = new Project(code, Value(row, map, "项目名称"), start, end);
        }
        return result;
    }

    private static Dictionary<string, List<string>> ReadPersonProjectMap(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var header = ws.RowsUsed().FirstOrDefault(r => r.CellsUsed().Any(c => CellText(c) == "姓名") &&
                                               r.CellsUsed().Any(c => Regex.IsMatch(CellText(c), "^RD\\d+", RegexOptions.IgnoreCase)));
        if (header is null)
        {
            header = ws.RowsUsed().FirstOrDefault(r => r.CellsUsed().Any(c => CellText(c) is "姓名" or "名称") &&
                                                    r.CellsUsed().Any(c => CellText(c) is "项目编号-外部" or "研发项目"))
                     ?? throw new InvalidDataException($"工作表“{ws.Name}”缺少人员和项目关系列");
            var longMap = HeaderMap(header);
            var longResult = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
            {
                var name = Value(row, longMap, "姓名");
                if (name.Length == 0) name = Value(row, longMap, "名称");
                var code = Value(row, longMap, "项目编号-外部");
                if (code.Length == 0) code = Value(row, longMap, "研发项目");
                if (name.Length == 0 || code.Length == 0) continue;
                if (!longResult.TryGetValue(name, out var codes)) longResult[name] = codes = [];
                if (!codes.Contains(code, StringComparer.OrdinalIgnoreCase)) codes.Add(code);
            }
            return longResult;
        }
        var headers = header.Cells(1, ws.LastColumnUsed()?.ColumnNumber() ?? 1)
            .ToDictionary(c => c.Address.ColumnNumber, CellText);
        var nameCol = headers.First(x => x.Value == "姓名").Key;
        var projectCols = headers.Where(x => Regex.IsMatch(x.Value, "^RD\\d+", RegexOptions.IgnoreCase)).ToArray();
        var result = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var row in ws.Rows(header.RowNumber() + 1, ws.LastRowUsed()?.RowNumber() ?? header.RowNumber()))
        {
            var name = CellText(row.Cell(nameCol));
            if (name.Length > 0) result[name] = projectCols.Where(x => CellText(row.Cell(x.Key)) == "是").Select(x => x.Value).ToList();
        }
        return result;
    }

    private static Dictionary<string, Dictionary<DateTime, string>> AssignProjects(
        List<DateTime> dates, List<Person> people,
        Dictionary<string, Dictionary<DateTime, string>> attendance,
        Dictionary<string, AttendanceRule> rules,
        Dictionary<string, Project> projects,
        Dictionary<string, List<string>> relations,
        HashSet<string> lockedMonths,
        AllocationOptions options,
        List<ProjectPeriodLock> periodLocks)
    {
        var result = new Dictionary<string, Dictionary<DateTime, string>>(StringComparer.Ordinal);
        foreach (var person in people)
        {
            var byDate = new Dictionary<DateTime, string>();
            string current = "";
            var blockCount = 0;
            var target = 0;
            DateTime? previousEffectiveDate = null;
            foreach (var date in dates)
            {
                var raw = attendance[person.Name].GetValueOrDefault(date, "");
                if (!rules.TryGetValue(raw, out var rule) || !rule.IsValid || rule.Days <= 0)
                {
                    byDate[date] = "";
                    if (options.InvalidBreaksBlock) { current = ""; blockCount = 0; }
                    continue;
                }
                if (lockedMonths.Contains(date.ToString("yyyy.MM")))
                {
                    byDate[date] = "";
                    continue;
                }
                var candidates = relations.GetValueOrDefault(person.Name, [])
                    .Where(code => projects.TryGetValue(code, out var p) &&
                                   (!p.Start.HasValue || p.Start.Value <= date) && (!p.End.HasValue || date < p.End.Value))
                    .Order(StringComparer.Ordinal).ToList();
                if (candidates.Count == 0) { byDate[date] = ""; continue; }

                var matchingLocks = periodLocks.Where(item =>
                    date >= item.Start && date <= item.End &&
                    ((item.EmployeeId.Length > 0 && item.EmployeeId == person.EmployeeId) ||
                     (item.Name.Length > 0 && item.Name == person.Name))).ToList();
                var lockedProjects = matchingLocks.Select(x => x.ProjectCode).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (lockedProjects.Count > 1) { byDate[date] = ""; continue; }
                if (lockedProjects.Count == 1 &&
                    (candidates.Contains(lockedProjects[0], StringComparer.OrdinalIgnoreCase) || matchingLocks.Any(x => x.Force)))
                {
                    current = lockedProjects[0];
                    byDate[date] = current;
                    blockCount++;
                    previousEffectiveDate = date;
                    continue;
                }

                if (!options.Enabled)
                {
                    byDate[date] = StablePick($"{person.EmployeeId}|{person.Name}|{date:yyyy-MM-dd}", candidates);
                    previousEffectiveDate = date;
                    continue;
                }

                var crossesMonth = previousEffectiveDate.HasValue && previousEffectiveDate.Value.Month != date.Month &&
                                   !options.ContinueAcrossMonths;
                var maximumReached = options.MaximumSameProjectDays.HasValue && blockCount >= options.MaximumSameProjectDays.Value;
                if (!candidates.Contains(current) || blockCount >= target || maximumReached || crossesMonth)
                {
                    var choices = options.AvoidFrequentSwitching && candidates.Count > 1
                        ? candidates.Where(x => x != current).ToList()
                        : candidates;
                    current = StablePick($"{person.EmployeeId}|{person.Name}|{date:yyyy-MM-dd}", choices);
                    target = options.MinimumDays + StableInt($"{person.Name}|{date:yyyy-MM-dd}|block",
                        options.MaximumDays - options.MinimumDays + 1);
                    blockCount = 0;
                }
                byDate[date] = current;
                blockCount++;
                previousEffectiveDate = date;
            }
            result[person.Name] = byDate;
        }
        return result;
    }

    private static void FillMatrixSheets(XLWorkbook workbook, List<DateTime> dates, List<Person> people,
        Dictionary<string, Dictionary<DateTime, string>> attendance,
        Dictionary<string, AttendanceRule> rules,
        Dictionary<string, Dictionary<DateTime, string>> assignments)
    {
        foreach (var sheetName in MatrixSheets)
        {
            var ws = workbook.Worksheet(sheetName);
            ResizeMatrix(ws, dates.Count, people.Count);
            for (var i = 0; i < dates.Count; i++)
            {
                var col = i + 5;
                ws.Cell(1, col).Value = dates[i];
                ws.Cell(1, col).Style.DateFormat.Format = "yyyy-mm-dd";
                ws.Cell(2, col).Value = "星期" + "日一二三四五六"[(int)dates[i].DayOfWeek];
            }
            for (var i = 0; i < people.Count; i++)
            {
                var row = i + 3;
                var person = people[i];
                ws.Cell(row, 1).Value = person.Seq;
                ws.Cell(row, 2).Value = person.EmployeeId;
                ws.Cell(row, 3).Value = person.Name;
                ws.Cell(row, 4).Value = person.Position;
                for (var d = 0; d < dates.Count; d++)
                {
                    var date = dates[d];
                    var raw = attendance[person.Name].GetValueOrDefault(date, "");
                    object value = sheetName switch
                    {
                        "全职人员考勤源表" => raw,
                        "全职人员考勤处理表" => rules.TryGetValue(raw, out var rule) ? rule.Normalized : raw,
                        "全职人员分项目情况表" => assignments[person.Name].GetValueOrDefault(date, ""),
                        _ => rules.TryGetValue(raw, out var hoursRule) ? hoursRule.Days : 0d
                    };
                    var target = ws.Cell(row, d + 5);
                    target.Clear(XLClearOptions.Contents);
                    if (value is not string text || text.Length > 0)
                        target.Value = XLCellValue.FromObject(value);
                }
            }
            ws.PageSetup.PrintAreas.Clear();
            ws.PageSetup.PrintAreas.Add(ws.Range(1, 1, Math.Max(3, people.Count + 2), Math.Max(4, dates.Count + 4)).RangeAddress.ToStringRelative());
        }
    }

    private static void ResizeMatrix(IXLWorksheet ws, int dateCount, int peopleCount)
    {
        var targetCols = Math.Max(5, dateCount + 4);
        var targetRows = Math.Max(3, peopleCount + 2);
        var currentCols = ws.LastColumnUsed()?.ColumnNumber() ?? 5;
        var currentRows = ws.LastRowUsed()?.RowNumber() ?? 3;
        if (currentCols < targetCols)
        {
            for (var col = currentCols + 1; col <= targetCols; col++)
            {
                ws.Column(col).Width = ws.Column(Math.Max(5, currentCols)).Width;
                for (var row = 1; row <= Math.Max(currentRows, targetRows); row++)
                    ws.Cell(row, col).Style = ws.Cell(row, Math.Max(5, currentCols)).Style;
            }
        }
        else if (currentCols > targetCols) ws.Columns(targetCols + 1, currentCols).Delete();
        if (currentRows < targetRows)
        {
            for (var row = currentRows + 1; row <= targetRows; row++)
            {
                ws.Row(row).Height = ws.Row(3).Height;
                for (var col = 1; col <= targetCols; col++) ws.Cell(row, col).Style = ws.Cell(3, col).Style;
            }
        }
        else if (currentRows > targetRows) ws.Rows(targetRows + 1, currentRows).Delete();
        foreach (var table in ws.Tables.ToArray()) table.Resize(ws.Range(1, 1, targetRows, targetCols));
    }

    private static void FillInternalReviewSheets(XLWorkbook wb, string rulePath, List<DateTime> dates, List<Person> people,
        Dictionary<string, Dictionary<DateTime, string>> attendance, Dictionary<string, AttendanceRule> rules,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects,
        Dictionary<string, List<string>> relations, HashSet<string> lockedMonths)
    {
        if (wb.TryGetWorksheet("主要问题提示", out var issues))
        {
            issues.Clear(XLClearOptions.Contents);
            issues.Cell(1, 1).InsertData(new[] { new[] { "序号", "问题类型", "人员", "日期", "说明" } });
            var rows = new List<string[]>();
            foreach (var person in people)
            foreach (var date in dates)
            {
                var raw = attendance[person.Name].GetValueOrDefault(date, "");
                if (rules.TryGetValue(raw, out var rule) && rule.IsValid && rule.Days > 0 &&
                    !lockedMonths.Contains(date.ToString("yyyy.MM")) &&
                    string.IsNullOrEmpty(assignments[person.Name].GetValueOrDefault(date, "")))
                    rows.Add(["", "有效考勤未分配项目", person.Name, date.ToString("yyyy-MM-dd"), "人员项目关系或项目周期未覆盖"]);
            }
            for (var i = 0; i < rows.Count; i++) rows[i][0] = (i + 1).ToString();
            if (rows.Count > 0) issues.Cell(2, 1).InsertData(rows);
            StyleReviewTable(issues.RangeUsed());
        }
        if (wb.TryGetWorksheet("系统复核表", out var review))
        {
            review.Clear(XLClearOptions.Contents);
            var headers = new[] { "工号或序号", "姓名", "日期", "原始考勤状态", "折算天数", "候选项目", "最终项目", "备注" };
            review.Cell(1, 1).InsertData(new[] { headers });
            var rows = new List<object[]>();
            foreach (var person in people)
            foreach (var date in dates)
            {
                var raw = attendance[person.Name].GetValueOrDefault(date, "");
                var candidates = relations.GetValueOrDefault(person.Name, []).Where(projects.ContainsKey);
                rows.Add([person.EmployeeId.Length > 0 ? person.EmployeeId : person.Seq, person.Name, date, raw,
                    rules.TryGetValue(raw, out var rule) ? rule.Days : 0d, string.Join(",", candidates),
                    assignments[person.Name].GetValueOrDefault(date, ""),
                    lockedMonths.Contains(date.ToString("yyyy.MM")) ? "历史月份锁定，未重新分配" : "按控制规则稳定分配"]);
            }
            if (rows.Count > 0) review.Cell(2, 1).InsertData(rows);
            review.Column(3).Style.DateFormat.Format = "yyyy-mm-dd";
            StyleReviewTable(review.RangeUsed());
        }
        if (wb.TryGetWorksheet("考勤规则拆解表", out var ruleSheet))
        {
            using var source = new XLWorkbook(rulePath);
            var used = source.Worksheets.First().RangeUsed();
            ruleSheet.Clear(XLClearOptions.Contents);
            if (used is not null) used.CopyTo(ruleSheet.Cell(1, 1));
        }
    }

    private int GenerateLogs(string outputRoot, string period, List<DateTime> dates, List<Person> people,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects,
        LogOptions options,
        IProgress<double>? progress)
    {
        var logRoot = Path.Combine(outputRoot, "3.01 研发工作日志");
        var daily = Path.Combine(logRoot, "日报");
        var weekly = Path.Combine(logRoot, "周报");
        var monthly = Path.Combine(logRoot, "月报");
        if (options.Daily) Directory.CreateDirectory(daily);
        if (options.Weekly) Directory.CreateDirectory(weekly);
        if (options.Monthly) Directory.CreateDirectory(monthly);
        var count = 0;
        for (var index = 0; index < people.Count; index++)
        {
            var person = people[index];
            if (options.Daily)
            {
                GenerateDailyLog(Path.Combine(daily, $"{Id(person)} {person.Name}研发工作日志-日报（{period}）.xlsx"), period, dates, person, assignments, projects);
                count++;
            }
            if (options.Weekly)
            {
                GenerateWeeklyLog(Path.Combine(weekly, $"{Id(person)} {person.Name}研发工作日志-周报（{period}）.xlsx"), dates, person, assignments, projects);
                count++;
            }
            if (options.Monthly)
            {
                GenerateMonthlyLog(Path.Combine(monthly, $"{Id(person)} {person.Name}研发工作日志-月报（{period}）.xlsx"), dates, person, assignments, projects);
                count++;
            }
            progress?.Report(65 + (index + 1) * 35d / people.Count);
        }
        return count;
    }

    private void GenerateDailyLog(string path, string period, List<DateTime> dates, Person person,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects)
    {
        using var wb = LoadTemplate("DailyLog.xlsx");
        var template = wb.Worksheets.First();
        foreach (var unused in wb.Worksheets.Skip(1).ToArray()) unused.Delete();
        var months = dates.Select(x => new DateTime(x.Year, x.Month, 1)).Distinct().Order().ToList();
        var created = new List<string>();
        foreach (var month in months)
        {
            var ws = created.Count == 0 ? template : template.CopyTo($"temp-{created.Count}");
            ws.Name = $"{month:yyyy.MM}工作日志";
            created.Add(ws.Name);
            var days = DateTime.DaysInMonth(month.Year, month.Month);
            ResizeLogRows(ws, days + 1);
            string previous = "";
            for (var day = 1; day <= days; day++)
            {
                var date = new DateTime(month.Year, month.Month, day);
                var row = day + 1;
                var code = assignments[person.Name].GetValueOrDefault(date, "");
                ws.Cell(row, 1).Value = day; ws.Cell(row, 2).Value = Id(person); ws.Cell(row, 3).Value = person.Name;
                ws.Cell(row, 4).Value = date; ws.Cell(row, 4).Style.DateFormat.Format = "yyyy-mm-dd";
                if (code.Length == 0) { ws.Range(row, 5, row, 7).Clear(XLClearOptions.Contents); continue; }
                var name = projects.GetValueOrDefault(code)?.Name ?? code;
                var text = MakeDailyText(code, name, date, previous);
                ws.Cell(row, 5).Value = code; ws.Cell(row, 6).Value = text; previous = text;
            }
        }
        wb.SaveAs(path);
    }

    private void GenerateWeeklyLog(string path, List<DateTime> dates, Person person,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects)
    {
        using var wb = LoadTemplate("WeeklyLog.xlsx");
        var baseSheet = wb.Worksheets.First();
        var start = dates.Min(); var end = dates.Max(); var cursor = start;
        var index = 0;
        while (cursor <= end)
        {
            var weekEnd = cursor.AddDays(6 - (((int)cursor.DayOfWeek + 6) % 7));
            if (weekEnd > end) weekEnd = end;
            var ws = index++ == 0 ? baseSheet : baseSheet.CopyTo($"周报{index}");
            ws.Name = $"{ISOWeek.GetYear(cursor)}年第{ISOWeek.GetWeekOfYear(cursor)}周工作日志";
            FillSummaryLog(ws, person, cursor, weekEnd, dates.Where(x => x >= cursor && x <= weekEnd), assignments, projects, "周");
            cursor = weekEnd.AddDays(1);
        }
        wb.SaveAs(path);
    }

    private void GenerateMonthlyLog(string path, List<DateTime> dates, Person person,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects)
    {
        using var wb = LoadTemplate("MonthlyLog.xlsx");
        var baseSheet = wb.Worksheets.First(); var index = 0;
        foreach (var month in dates.GroupBy(x => (x.Year, x.Month)).OrderBy(x => x.Key))
        {
            var ws = index++ == 0 ? baseSheet : baseSheet.CopyTo($"月报{index}");
            ws.Name = $"{month.Key.Year}年{month.Key.Month:00}月工作日志";
            FillSummaryLog(ws, person, month.Min(), month.Max(), month, assignments, projects, "月");
        }
        wb.SaveAs(path);
    }

    private static void FillSummaryLog(IXLWorksheet ws, Person person, DateTime start, DateTime end, IEnumerable<DateTime> dates,
        Dictionary<string, Dictionary<DateTime, string>> assignments, Dictionary<string, Project> projects, string unit)
    {
        var codes = dates.Select(x => assignments[person.Name].GetValueOrDefault(x, "")).Where(x => x.Length > 0).Distinct().ToList();
        var names = codes.Select(x => projects.GetValueOrDefault(x)?.Name ?? x).ToList();
        var joined = names.Count == 0 ? "本阶段研发工作" : string.Join("、", names);
        ws.Cell("B2").Value = person.Name; ws.Cell("D2").Value = Id(person); ws.Cell("F2").Value = $"{start:yyyy/MM/dd}-{end:yyyy/MM/dd}";
        ws.Cell("A4").Value = $"本{unit}围绕{joined}推进研发工作，完成技术资料梳理、研发过程复核和阶段事项跟进。";
        ws.Cell("A6").Value = $"结合本{unit}推进情况，重点关注技术资料完整性、工序衔接和待验证事项。";
        ws.Cell("A8").Value = $"下{unit}将延续当前研发阶段，继续核对关键技术环节并整理后续验证事项。";
    }

    private static void ResizeLogRows(IXLWorksheet ws, int targetRows)
    {
        var current = ws.LastRowUsed()?.RowNumber() ?? 2;
        if (current > targetRows) ws.Rows(targetRows + 1, current).Delete();
        else for (var row = current + 1; row <= targetRows; row++)
        {
            ws.Row(row).Height = ws.Row(2).Height;
            for (var col = 1; col <= (ws.LastColumnUsed()?.ColumnNumber() ?? 7); col++) ws.Cell(row, col).Style = ws.Cell(2, col).Style;
        }
    }

    private static string MakeDailyText(string code, string projectName, DateTime date, string previous)
    {
        var actions = new[] { "梳理", "复核", "比对", "分析" };
        var objects = new[] { "工艺路线", "技术参数", "质量控制", "阶段验证事项" };
        var action = actions[StableInt($"{code}|{date:yyyy-MM-dd}|a", actions.Length)];
        var obj = objects[StableInt($"{code}|{date:yyyy-MM-dd}|o", objects.Length)];
        return previous.Length == 0
            ? $"结合现有记录，{action}{projectName}的{obj}，同步整理阶段研发要点，供后续验证参考。"
            : $"延续前期记录，继续{action}{projectName}相关的{obj}，列出仍需确认的技术环节，准备后续复核。";
    }

    private static XLWorkbook LoadTemplate(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resource = assembly.GetManifestResourceNames().FirstOrDefault(x => x.EndsWith($"Templates.{fileName}", StringComparison.Ordinal));
        if (resource is null) throw new InvalidOperationException($"应用内置模板缺失：{fileName}");
        using var stream = assembly.GetManifestResourceStream(resource)!;
        var memory = new MemoryStream(); stream.CopyTo(memory); memory.Position = 0;
        return new XLWorkbook(memory);
    }

    private static (string Company, string Period, string GeneratedOn) DetectMetadata(IEnumerable<string> files, List<DateTime> dates)
    {
        var text = string.Join(" ", files.Select(Path.GetFileName));
        var company = Regex.Match(text, @"\b[A-Z]{2,8}\b").Value;
        if (company.Length == 0)
        {
            company = files.Select(Path.GetFileNameWithoutExtension)
                .Select(name => Regex.Match(name ?? "", @"^\d+\.\d+\s+(.+?)(?=20\d{2}\.\d{2})"))
                .FirstOrDefault(match => match.Success)?.Groups[1].Value.Trim() ?? "";
        }
        var period = Regex.Match(text, @"20\d{2}\.\d{2}-20\d{2}\.\d{2}").Value;
        return (company.Length == 0 ? "公司" : company,
            period.Length == 0 ? $"{dates.Min():yyyy.MM}-{dates.Max():yyyy.MM}" : period,
            DateTime.Today.ToString("yyyyMMdd"));
    }

    private static (DateTime? Start, DateTime? End) ParsePeriod(string text)
    {
        var fullDates = Regex.Match(text, @"(20\d{2})[./-](\d{1,2})[./-](\d{1,2})\s*[-至]\s*(20\d{2})[./-](\d{1,2})[./-](\d{1,2})");
        if (fullDates.Success)
        {
            var startDate = new DateTime(int.Parse(fullDates.Groups[1].Value), int.Parse(fullDates.Groups[2].Value), int.Parse(fullDates.Groups[3].Value));
            var endDate = new DateTime(int.Parse(fullDates.Groups[4].Value), int.Parse(fullDates.Groups[5].Value), int.Parse(fullDates.Groups[6].Value));
            return (startDate, endDate.AddDays(1));
        }
        var match = Regex.Match(text, @"(20\d{2})\.(\d{2})-(20\d{2})\.(\d{2})");
        if (!match.Success) return (null, null);
        var start = new DateTime(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value), 1);
        var endMonth = int.Parse(match.Groups[4].Value); var endYear = int.Parse(match.Groups[3].Value);
        return (start, new DateTime(endYear, endMonth, 1).AddMonths(1));
    }

    private static IXLRow FindHeaderRow(IXLWorksheet ws, params string[] required) =>
        ws.RowsUsed().FirstOrDefault(row => required.All(key => row.CellsUsed().Any(c => CellText(c) == key)))
        ?? throw new InvalidDataException($"工作表“{ws.Name}”缺少标题列：{string.Join("、", required)}");

    private static Dictionary<string, int> HeaderMap(IXLRow row) => row.CellsUsed()
        .Where(c => CellText(c).Length > 0).GroupBy(CellText).ToDictionary(g => g.Key, g => g.First().Address.ColumnNumber, StringComparer.Ordinal);
    private static string Value(IXLRow row, IReadOnlyDictionary<string, int> map, string key) => map.TryGetValue(key, out var col) ? CellText(row.Cell(col)) : "";
    private static string CellText(IXLCell cell) => cell.GetFormattedString().Replace("　", " ").Trim();
    private static bool IsDateCell(IXLCell cell) => ReadDate(cell).HasValue;
    private static DateTime? ReadDate(IXLCell cell)
    {
        if (cell.TryGetValue<DateTime>(out var date) && date.Year is >= 2000 and <= 2100) return date;
        var text = cell.GetFormattedString().Trim();
        foreach (var format in new[] { "yyyy-MM-dd", "yyyy/M/d", "yyyy.MM.dd", "M/d/yyyy" })
            if (DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Year >= 2000) return date;
        return null;
    }
    private static string StablePick(string seed, List<string> values) => values[StableInt(seed, values.Count)];
    private static int StableInt(string seed, int max)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(seed));
        return max <= 1 ? 0 : (int)(BitConverter.ToUInt32(hash, 0) % max);
    }
    private static string Id(Person person) => person.EmployeeId.Length > 0 ? person.EmployeeId : person.Seq;
    private static string? NullIfEmpty(string value) => string.IsNullOrWhiteSpace(value) ? null : value;
    private static int PositiveInt(string? value, int fallback) =>
        int.TryParse(value, out var parsed) && parsed > 0 ? parsed : fallback;
    private static int? NullablePositiveInt(string? value) =>
        int.TryParse(value, out var parsed) && parsed > 0 ? parsed : null;
    private static string NormalizeMonth(string value)
    {
        var match = Regex.Match(value, @"(20\d{2})[./-](\d{1,2})");
        return match.Success ? $"{match.Groups[1].Value}.{int.Parse(match.Groups[2].Value):00}" : "";
    }
    private static void StyleReviewTable(IXLRange? range)
    {
        if (range is null) return;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        range.Style.Alignment.WrapText = true;
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        range.FirstRow().Style.Font.Bold = true;
        range.FirstRow().Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAF7");
        range.Worksheet.Columns(1, range.ColumnCount()).AdjustToContents(8, 36);
    }
}
