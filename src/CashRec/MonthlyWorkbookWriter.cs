using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace CashRec;

internal static class MonthlyWorkbookWriter
{
    const string AmountFormat = "\"$\"#,##0.00;[Red]\\-\"$\"#,##0.00";
    const string NumberFormat = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

    public static void Write(string temporaryPath, string? existingPath, MonthlyReport report)
    {
        using var book = existingPath == null ? NewWorkbook() : new XLWorkbook(existingPath);
        Validate(book);
        SetProperty(book, "CashRecMonthlyStage", MonthlyStage.Current);
        SetProperty(book, "CashRecMonthlyStageComplete", report.StageComplete ? "1" : "0");
        SetProperty(book, "CashRecVersion", AppVersion.Current);
        var rows = MonthlyLayout.Sheets.ToDictionary(name => name, name => FindOrAddDate(book.Worksheet(name), report.Date));
        var caches = new Dictionary<string, Dictionary<string, decimal>>();
        foreach (var sheet in book.Worksheets)
        {
            var cached = new Dictionary<string, decimal>();
            foreach (var cell in sheet.CellsUsed().Where(cell => cell.HasFormula && cell.CachedValue.Type == XLDataType.Number))
                cached[Address(cell)] = ExcelNumbers.Read(cell, sheet.Name + "!" + cell.Address);
            caches["xl/worksheets/sheet" + sheet.Position + ".xml"] = cached;
        }
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.Worksheet(name);
            var columns = name == "Break down" ? MonthlyLayout.BreakdownColumns : MonthlyLayout.EntityColumns;
            var number = rows[name];
            var values = report.Rows[name];
            for (var column = 0; column < columns.Length; column++)
            {
                var field = columns[column].Field;
                var cell = sheet.Cell(number, column + 1);
                // A missing previous value must also remove a previously saved DIFF.
                if (name != "Break down" && IsAutomaticField(field))
                {
                    cell.Clear(XLClearOptions.Contents);
                    caches["xl/worksheets/sheet" + sheet.Position + ".xml"].Remove(Address(cell));
                }
                if (!values.TryGetValue(field, out var value) || value == null) continue;
                switch (value)
                {
                    case DateTime date: cell.Value = date; cell.Style.DateFormat.Format = "mm-dd-yy"; break;
                    case decimal amount: cell.Value = (double)amount; break;
                    case MonthlyFormula formula:
                        cell.FormulaA1 = Resolve(formula.Expression, name, rows);
                        caches["xl/worksheets/sheet" + sheet.Position + ".xml"][Address(cell)] = formula.CalculatedValue;
                        break;
                    default: throw new InvalidDataException("不支持的月度字段值：" + field);
                }
            }
        }
        book.SaveAs(temporaryPath, new SaveOptions { EvaluateFormulasBeforeSaving = false });
        WorkbookWriter.PersistFormulaValues(temporaryPath, caches);
    }

    static bool IsAutomaticField(string field) => field == "SETTLE PNL"
        || MonthlyFlowCalculator.Fields.Contains(field)
        || MonthlyBalanceCalculator.BalanceFields.Contains(field)
        || MonthlyBalanceCalculator.BalanceFields.Any(balance => field == balance + " DIFF");

    static string Address(IXLCell cell) => WorkbookWriter.CellAddress(cell.Address.ColumnNumber, cell.Address.RowNumber);

    static string Resolve(string expression, string sheet, IReadOnlyDictionary<string, int> rows) =>
        Regex.Replace(expression, @"\[([^\]]+)\]", match =>
        {
            var parts = match.Groups[1].Value.Split(':', 2);
            var target = parts.Length == 2 ? parts[0] : sheet;
            var field = parts[^1];
            var address = MonthlyLayout.Address(target, field, rows[target]);
            return target == sheet ? address : "'" + target.Replace("'", "''") + "'!" + address;
        });

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

    static void Validate(XLWorkbook book)
    {
        if (book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecMonthlyRules")?.Value.ToString() != MonthlyLayout.Rules
            || book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecRules")?.Value.ToString() != ReportPolicy.EntityRules
            || book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecValuation")?.Value.ToString() != ReportPolicy.Valuation)
            throw new InvalidDataException("已有月度工作簿口径或布局无法确认，已保留原文件；请提供当前版本月度文件");
        if (!book.Worksheets.Select(sheet => sheet.Name).SequenceEqual(MonthlyLayout.Sheets))
            throw new InvalidDataException("已有月度工作簿的四个工作表不完整或顺序不符");
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.Worksheet(name);
            var columns = name == "Break down" ? MonthlyLayout.BreakdownColumns : MonthlyLayout.EntityColumns;
            for (var c = 0; c < columns.Length; c++)
                if (sheet.Cell(MonthlyLayout.HeaderRows(name), c + 1).GetString() != columns[c].Label)
                    throw new InvalidDataException(name + " 月度表头与当前字段布局不一致：" + columns[c].Label);
        }
    }

    static int FindOrAddDate(IXLWorksheet sheet, DateTime date)
    {
        var headerRows = MonthlyLayout.HeaderRows(sheet.Name);
        var matches = sheet.RowsUsed().Where(row => row.RowNumber() > headerRows && CellDate(row.Cell(1)) == date).ToList();
        if (matches.Count > 1) throw new InvalidDataException(sheet.Name + " 同一业务日期重复，不能任选一行覆盖");
        if (matches.Count == 1) return matches[0].RowNumber();
        var last = sheet.RowsUsed().Where(row => row.RowNumber() > headerRows && CellDate(row.Cell(1)) != null)
            .Select(row => row.RowNumber()).DefaultIfEmpty(headerRows).Max();
        var next = last + 1;
        if (last > headerRows) sheet.Row(next).Style = sheet.Row(last).Style;
        if (last > headerRows)
            foreach (var cell in sheet.Row(last).CellsUsed(XLCellsUsedOptions.NormalFormats))
                sheet.Cell(next, cell.Address.ColumnNumber).Style = cell.Style;
        return next;
    }

    static DateTime? CellDate(IXLCell cell)
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
