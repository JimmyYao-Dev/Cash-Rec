using System.Text.RegularExpressions;

namespace CashRec;

/// <summary>Current-day settlement, signed commission and the established manual Check.</summary>
public static class MonthlySettlementCalculator
{
    public static readonly string[] ManualFields = ["To LP", "Bank Fee", "Bank fee-conversion", "Time Diff"];
    public static readonly string[] AutomaticFields = ["TMA SETTLE PNL", "TMG SETTLE PNL", "To Settle WS", "Commission", "Check"];

    public static void Apply(MonthlyReport report, IEnumerable<ReportTable> tables,
        IReadOnlyDictionary<string, decimal?>? manual = null, Action<string>? progress = null)
    {
        var row = report.Rows["Break down"];
        var tma = MonthlyBalanceCalculator.Number(report.Rows["TMA"].GetValueOrDefault("SETTLE PNL"), "TMA SETTLE PNL");
        var tmg = MonthlyBalanceCalculator.Number(report.Rows["TMG"].GetValueOrDefault("SETTLE PNL"), "TMG SETTLE PNL");
        row["TMA SETTLE PNL"] = new MonthlyFormula("[TMA:SETTLE PNL]", tma);
        row["TMG SETTLE PNL"] = new MonthlyFormula("[TMG:SETTLE PNL]", tmg);
        row["To Settle WS"] = 0m;
        var commission = 0m;
        var count = 0;
        var processed = 0;
        foreach (var transaction in tables.Where(table => table.Name == "TransactionExchange").SelectMany(table => table.Rows))
        {
            if (++processed % 10000 == 0) progress?.Invoke($"汇总结算佣金：{processed:N0} 条准备记录");
            var type = Text(transaction, "Transaction Type");
            if (type is not ("PSP CN COMMISSION" or "PSP COMMISSION") || Text(transaction, "Branch") is not ("TMA" or "TMG")) continue;
            if (transaction.GetValueOrDefault("Date") is not DateTime date) throw new InvalidDataException("佣金记录缺少有效报表归属日期");
            if (date.Date != report.Date) continue;
            commission += MonthlyBalanceCalculator.Number(transaction.GetValueOrDefault("USD Equivalent"), "佣金 USD Equivalent");
            count++;
        }
        row["Commission"] = commission;
        if (!row.TryGetValue("Diff", out var difference) || difference == null)
        {
            row["Check"] = null;
            report.CheckComplete = false;
            report.Diagnostics.Add(new("Incomplete", "Check", "资金差额缺失，Check留空；未将缺失差额当零，结算与佣金仍保留当前可用值"));
            return;
        }
        var result = MonthlyBalanceCalculator.Number(difference, "Break down Diff") - tma - tmg - commission
            - ManualFields.Sum(field => manual?.GetValueOrDefault(field) ?? 0m);
        row["Check"] = new MonthlyFormula("[Diff]-[To LP]-[TMA SETTLE PNL]-[TMG SETTLE PNL]-[To Settle WS]-[Commission]-[Bank Fee]-[Bank fee-conversion]-[Time Diff]", result);
        report.CheckComplete = true;
        report.Provenance.Add(new
        {
            Kind = "Settlement and Check", Date = report.Date,
            Tma = "Current report TMA SETTLE PNL, no TMNZ", Tmg = "Current report TMG SETTLE PNL",
            CommissionRecords = count, Commission = "Same report date, TMA/TMG, exact normalized PSP CN Commission or PSP Commission; all Sources; signed USD Equivalent",
            Check = "F-G-H-I-J-K-L-M-N, preserving manual blanks/zero/negative values; no ROUND or tolerance",
        });
    }

    static string Text(IReadOnlyDictionary<string, object?> row, string field) =>
        Regex.Replace(row.GetValueOrDefault(field)?.ToString()?.Trim() ?? "", @"\s+", " ").ToUpperInvariant();
}
