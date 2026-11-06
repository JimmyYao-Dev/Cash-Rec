using System.Globalization;
using System.Security.Cryptography;
using ClosedXML.Excel;

namespace CashRec;

/// <summary>Resolves a read-only, same-policy month-end source for a new monthly file.</summary>
internal sealed class MonthlyRollover(string path, DateTime date, FileStream guard) : IDisposable
{
    internal string Path { get; } = path;
    internal DateTime Date { get; } = date;
    public void Dispose() => guard.Dispose();

    internal static MonthlyRollover? Resolve(RunContext context, string destination, MonthlyReport report)
    {
        if (context.PreviousBalances == null) return null;
        var previous = context.PreviousBalances.Date;
        var month = new DateTime(context.BusinessDate.Year, context.BusinessDate.Month, 1);
        var priorMonth = month.AddMonths(-1);
        var awaitingPrevious = false;
        if (File.Exists(destination))
        {
            using var existing = new XLWorkbook(destination);
            awaitingPrevious = existing.CustomProperties.SingleOrDefault(property => property.Name == "CashRecMonthlyBaselineState")?.Value.ToString() == "AwaitingPrevious";
        }
        if (previous >= month && !awaitingPrevious) return null; // A known same-month basis permits mid-month bootstrap.
        if (previous < priorMonth)
            throw new InvalidDataException($"跨月缺少上月有效基准；上期余额日期为{previous:yyyy.MM.dd}，请先通过界面生成上月最后有效业务日");
        if (File.Exists(destination))
        {
            using var current = new XLWorkbook(destination);
            MonthlyWorkbookWriter.Validate(current);
            var carried = MonthlyLayout.Sheets.Select(name => current.Worksheet(name).RowsUsed(XLCellsUsedOptions.Contents)
                .Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(name) && !row.Cell(1).HasFormula)
                .Select(row => MonthlyWorkbookWriter.CellDate(row.Cell(1))).Where(date => date < month).ToList()).ToList();
            if (carried.All(dates => dates.Count == 1 && dates[0] == previous)) return null;
            if (carried.Any(dates => dates.Count > 0))
                throw new InvalidDataException("本月四表已有的跨月基准日期不一致或不等于当前上期余额日；已保留现有文件和人工值");
            // An earlier no-prior partial output is not an initialized rollover.
            // Add the missing carry row without clearing any current-month dates.
        }
        var source = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(destination)!, $"CashRec_Breakdown_{priorMonth:yyyy.MM}.xlsx");
        if (!File.Exists(source))
            throw new InvalidDataException($"缺少上月月度文件：{source}；请先通过界面生成上月最后有效业务日，再生成本月");
        var held = OutputTransaction.Guard(source);
        try
        {
            using var book = new XLWorkbook(source);
            MonthlyWorkbookWriter.Validate(book);
            var dates = MonthlyLayout.Sheets.Select(name => LastDate(book.Worksheet(name), priorMonth, month)).ToList();
            if (dates.Distinct().Count() != 1)
                throw new InvalidDataException("上月四表最后有效业务日期不一致：" + string.Join("；", MonthlyLayout.Sheets.Zip(dates, (name, date) => name + "=" + date.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture))));
            var baseline = dates[0];
            if (previous < month && baseline != previous)
                throw new InvalidDataException($"上月四表最后有效日期{baseline:yyyy.MM.dd}与当前所用上期余额日{previous:yyyy.MM.dd}不一致；请先通过界面生成对应上月基准");
            if (!BalanceWorkbookReader.HasCurrentClassification(source, baseline))
                throw new InvalidDataException($"请先按新Comment规则重跑上月最后有效日{baseline:yyyy.MM.dd}；原月度文件及人工值已保留");
            ValidateBalances(book, baseline);
            using var input = File.OpenRead(source);
            report.Provenance.Add(new { Kind = "Monthly rollover", BaselineDate = baseline, SourcePath = source,
                SourceSHA256 = Convert.ToHexString(SHA256.HashData(input)), Selection = "Same last effective date in all four sheets; previous month source is read-only" });
            report.Diagnostics.Add(new("Info", "跨月基准", $"新月仅携带上月最后有效日{baseline:yyyy.MM.dd}的四表记录及人工值；上月文件保持原样"));
            return new(source, baseline, held);
        }
        catch { held.Dispose(); throw; }
    }

    static DateTime LastDate(IXLWorksheet sheet, DateTime from, DateTime until)
    {
        var dates = sheet.RowsUsed(XLCellsUsedOptions.Contents).Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(sheet.Name))
            .Select(row => row.Cell(1)).Where(cell => !cell.HasFormula).Select(MonthlyWorkbookWriter.CellDate)
            .Where(date => date >= from && date < until).Select(date => date!.Value).ToList();
        if (dates.Count == 0) throw new InvalidDataException(sheet.Name + " 上月没有有效业务日期，不能用空白、总计或说明行替代");
        var last = dates.Max();
        if (dates.Count(date => date == last) != 1) throw new InvalidDataException(sheet.Name + " 上月最后业务日期重复，不能任选基准");
        return last;
    }

    static void ValidateBalances(XLWorkbook book, DateTime date)
    {
        foreach (var name in new[] { "TMA", "TMG", "Total" })
        {
            var sheet = book.Worksheet(name);
            var row = sheet.RowsUsed().Single(row => row.RowNumber() > MonthlyLayout.HeaderRows(name) && !row.Cell(1).HasFormula && MonthlyWorkbookWriter.CellDate(row.Cell(1)) == date);
            foreach (var field in MonthlyBalanceCalculator.BalanceFields.Concat(["SETTLE PNL"]))
            {
                var column = Array.FindIndex(MonthlyLayout.EntityColumns, candidate => candidate.Field == field) + 1;
                ExcelNumbers.Read(row.Cell(column), $"{name} {date:yyyy.MM.dd} 上月基准 {field}");
            }
        }
    }
}
