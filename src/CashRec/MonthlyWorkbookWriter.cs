using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace CashRec;

internal static class MonthlyWorkbookWriter
{
    const string AmountFormat = "\"$\"#,##0.00;[Red]\\-\"$\"#,##0.00";
    const string NumberFormat = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

    public static void Write(string temporaryPath, string? existingPath, MonthlyReport report, MonthlyRollover? rollover = null)
    {
        var source = existingPath ?? rollover?.Path;
        using var book = source == null ? NewWorkbook() : new XLWorkbook(source);
        using var additional = existingPath != null && rollover != null ? new XLWorkbook(rollover.Path) : null;
        Validate(book);
        for (var index = 0; index < MonthlyLayout.Sheets.Length; index++) book.Worksheet(MonthlyLayout.Sheets[index]).Position = index + 1;
        SetProperty(book, "CashRecMonthlyStage", MonthlyStage.Current);
        SetProperty(book, "CashRecMonthlyStageComplete", report.StageComplete ? "1" : "0");
        SetProperty(book, "CashRecMonthlyStageDate", report.Date.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture));
        SetProperty(book, "CashRecVersion", AppVersion.Current);
        var caches = MonthlyHistory.Rewrite(book, report, existingPath == null ? rollover?.Date : null, additional, rollover?.Date);
        if (rollover != null)
        {
            SetProperty(book, "CashRecMonthlyBaselineDate", rollover.Date.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture));
            SetProperty(book, "CashRecMonthlyBaselineState", "Initialized");
        }
        else if (source == null)
            SetProperty(book, "CashRecMonthlyBaselineState", report.PreviousDate == null
                ? report.Date.Day == 1 ? "AwaitingPrevious" : "SameMonthBootstrapNoPrior"
                : "SameMonthBootstrap");
        MarkClassificationCoverage(book, report);
        book.SaveAs(temporaryPath, new SaveOptions { EvaluateFormulasBeforeSaving = false });
        WorkbookWriter.PersistFormulaValues(temporaryPath, caches);
    }

    internal static Dictionary<string, decimal?> ReadManualInputs(string? path, DateTime date)
    {
        var result = MonthlySettlementCalculator.ManualFields.ToDictionary(field => field, _ => (decimal?)null);
        if (path == null) return result;
        using var book = new XLWorkbook(path);
        Validate(book);
        var sheet = book.Worksheet("Break down");
        var rows = sheet.RowsUsed().Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(sheet.Name) && CellDate(row.Cell(1)) == date).ToList();
        if (rows.Count > 1) throw new InvalidDataException("Break down 同日人工区域存在重复业务日期，不能任选一行");
        if (rows.Count == 0) return result;
        foreach (var field in MonthlySettlementCalculator.ManualFields)
        {
            var column = Array.FindIndex(MonthlyLayout.BreakdownColumns, candidate => candidate.Field == field) + 1;
            var cell = rows[0].Cell(column);
            if (!cell.HasFormula && cell.Value.Type == XLDataType.Blank) continue;
            result[field] = ExcelNumbers.Read(cell, $"Break down {date:yyyy.MM.dd} {field}");
        }
        return result;
    }

    internal static bool IsAutomaticField(string field) => field == "SETTLE PNL"
        || field is "Insto Pending DIFF" or "Retail Pending DIFF" or "Wholesale Pending DIFF"
        || MonthlyFlowCalculator.Fields.Contains(field)
        || MonthlyBalanceCalculator.BalanceFields.Contains(field)
        || MonthlyBalanceCalculator.BalanceFields.Any(balance => field == balance + " DIFF");

    static string Address(IXLCell cell) => WorkbookWriter.CellAddress(cell.Address.ColumnNumber, cell.Address.RowNumber);

    static void MarkClassificationCoverage(XLWorkbook book, MonthlyReport report)
    {
        SetProperty(book, ReportPolicy.ClassificationDateProperty(report.Date), ReportPolicy.Classification);
        var dates = book.Worksheets.SelectMany(sheet => sheet.RowsUsed()
            .Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(sheet.Name))
            .Select(row => CellDate(row.Cell(1))))
            .Where(date => date != null).Select(date => date!.Value).Distinct().Order().ToList();
        var stale = dates.Where(date => book.CustomProperties.SingleOrDefault(property =>
                property.Name == ReportPolicy.ClassificationDateProperty(date))?.Value.ToString() != ReportPolicy.Classification)
            .ToList();
        SetProperty(book, "CashRecClassification", stale.Count == 0 ? ReportPolicy.Classification : ReportPolicy.MixedClassification);
        if (stale.Count > 0)
            report.Diagnostics.Add(new("Warning", "历史交易分类",
                $"{stale.Count} 个历史日期尚未按当前 Comment 关键词规则重算（{string.Join("、", stale.Take(5).Select(date => date.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture)))}）；已保留原行及人工值，月度分类标记为 Mixed，请逐日重跑，不将旧自动值作为新分类结果"));
    }

    static XLWorkbook NewWorkbook()
    {
        var book = new XLWorkbook();
        book.CustomProperties.Add("CashRecMonthlyRules", MonthlyLayout.Rules);
        book.CustomProperties.Add("CashRecRules", ReportPolicy.EntityRules);
        book.CustomProperties.Add("CashRecValuation", ReportPolicy.Valuation);
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.AddWorksheet(name);
            var columns = name == "Break down" ? MonthlyLayout.BreakdownColumns : MonthlyLayout.EntityColumns;
            var headerRows = MonthlyLayout.HeaderRows(name);
            for (var c = 0; c < columns.Length; c++)
            {
                sheet.Cell(headerRows, c + 1).Value = columns[c].Label;
                sheet.Column(c + 1).Width = columns[c].Field.Length == 0 ? 3 : 19;
                var format = c == 0 ? "mm-dd-yy" : columns[c].Field.Length == 0 ? "General" : AmountFormat;
                if (name == "Break down")
                    format = c == 0 ? "mm-dd-yy" : c is 13 or 14 ? "General"
                        : c is 2 or 3 or 5 or 11 or 12 or 15 ? NumberFormat : AmountFormat;
                sheet.Cell(headerRows + 1, c + 1).Style.NumberFormat.Format = format;
            }
            if (name != "Break down")
            {
                foreach (var group in new[] { (1, 5, "MT4"), (8, 13, "BANK"), (16, 25, "DIFF") })
                {
                    sheet.Range(1, group.Item1, 1, group.Item2).Merge();
                    sheet.Cell(1, group.Item1).Value = group.Item3;
                }
            }
            sheet.Range(1, 1, headerRows, columns.Length).Style.Font.Bold = true;
            sheet.Range(1, 1, headerRows, columns.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("D9E5EF");
            sheet.SheetView.FreezeRows(headerRows);
        }
        return book;
    }

    internal static void Validate(XLWorkbook book)
    {
        if (book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecMonthlyRules")?.Value.ToString() != MonthlyLayout.Rules
            || book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecRules")?.Value.ToString() != ReportPolicy.EntityRules
            || book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecValuation")?.Value.ToString() != ReportPolicy.Valuation)
            throw new InvalidDataException("已有月度工作簿口径或布局无法确认，已保留原文件；请提供当前版本月度文件");
        if (book.Worksheets.Count != MonthlyLayout.Sheets.Length || !book.Worksheets.Select(sheet => sheet.Name).ToHashSet(StringComparer.Ordinal).SetEquals(MonthlyLayout.Sheets))
            throw new InvalidDataException("已有月度工作簿的四个工作表不完整");
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.Worksheet(name);
            var columns = name == "Break down" ? MonthlyLayout.BreakdownColumns : MonthlyLayout.EntityColumns;
            for (var c = 0; c < columns.Length; c++)
                if (sheet.Cell(MonthlyLayout.HeaderRows(name), c + 1).GetString() != columns[c].Label)
                    throw new InvalidDataException(name + " 月度表头与当前字段布局不一致：" + columns[c].Label);
        }
    }

    internal static DateTime? CellDate(IXLCell cell)
    {
        if (cell.HasFormula) throw new InvalidDataException(cell.Worksheet.Name + " 日期不能由未验证公式推算");
        if (cell.Value.Type == XLDataType.Blank) return null;
        if (cell.TryGetValue<DateTime>(out var date)) return date.Date;
        return DateTime.TryParseExact(cell.GetString(), ["yyyy.MM.dd", "yyyy-MM-dd"], CultureInfo.InvariantCulture,
            DateTimeStyles.None, out date) ? date.Date : null;
    }

    static void SetProperty(XLWorkbook book, string name, string value)
    {
        var existing = book.CustomProperties.SingleOrDefault(property => property.Name == name);
        if (existing == null) book.CustomProperties.Add(name, value);
        else existing.Value = value;
    }
}
