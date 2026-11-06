using System.Globalization;
using System.Text.RegularExpressions;

namespace CashRec;

internal static class MonthlyBalanceCalculator
{
    internal static readonly string[] BalanceFields = ["Insto", "Retail", "Wholesale", "PSP", "PSP Pending", "Settle"];
    static readonly HashSet<string> TmaSettleTypes = new(StringComparer.Ordinal)
    { "SETTLE", "POSTING TF TD", "MISC CIMB TF TD", "CFD TF TD", "REBATE TF TD", "PENDING TF TD" };

    public static MonthlyReport Calculate(RunContext context, IList<ReportTable> tables)
    {
        var report = new MonthlyReport(context.BusinessDate) { PreviousDate = context.PreviousBalances?.Date };
        var bank = tables.Single(table => table.Name == "Bank").Rows;
        var psp = tables.Single(table => table.Name == "PSP").Rows;
        var transactions = tables.Where(table => table.Name == "TransactionExchange").SelectMany(table => table.Rows).ToList();
        var pending = PendingSnapshotReader.Read(context.SourcePath("BankBD"), context.BusinessDate);
        report.Provenance.Add(new { Kind = "PSP Pending", Snapshot = pending });
        PendingSnapshot? previousPending = null;
        var previous = context.PreviousBalances;
        if (previous == null)
            report.Diagnostics.Add(new("Incomplete", "月度余额", "没有有效同口径上期；当前余额保留，DIFF留空，不当零", DiagnosticCode.MissingPriorBasis));
        else
        {
            report.Provenance.Add(new { Kind = "Previous balances", previous.Date, previous.Source, previous.DateDirectory });
            try
            {
                var priorPath = Path.Combine(previous.DateDirectory, "Needed files", "PSP+BANK", "BANK BD.xlsx");
                previousPending = PendingSnapshotReader.Read(priorPath, previous.Date);
                report.Provenance.Add(new { Kind = "Previous PSP Pending", Snapshot = previousPending });
            }
            catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException)
            {
                report.Diagnostics.Add(new("Incomplete", "月度PSP Pending",
                    $"上期 {previous.Date:yyyy.MM.dd} 的 BANK BD 快照不可用：{error.Message}；PSP Pending DIFF留空，不借用当前累计值", DiagnosticCode.MissingPriorPending));
            }
        }
        foreach (var branch in new[] { "TMA", "TMG" })
        {
            var row = report.Rows[branch];
            row["Date"] = context.BusinessDate;
            var current = Aggregate(bank, psp, branch);
            current["PSP Pending"] = pending.Entities[branch].Pending;
            foreach (var (field, value) in current) row[field] = value;
            var retail = transactions.Where(transaction => Text(transaction, "Branch") == branch
                && Text(transaction, "Type") == "RETAIL" && Organization(transaction) == "TMC"
                && transaction.GetValueOrDefault("Date") is DateTime reportDate && reportDate.Date == context.BusinessDate).ToList();
            row["SETTLE PNL"] = retail.Where(transaction => IsSettleType(branch, Text(transaction, "Transaction Type")))
                .Sum(transaction => Number(transaction.GetValueOrDefault("USD Equivalent"), "SETTLE PNL USD Equivalent"));
            if (previous == null) continue;
            var prior = Aggregate(previous.Tables["Bank"], previous.Tables["PSP"], branch);
            if (previousPending != null) prior["PSP Pending"] = previousPending.Entities[branch].Pending;
            foreach (var field in BalanceFields)
            {
                if (!prior.TryGetValue(field, out var value)) continue;
                row[field + " DIFF"] = new MonthlyFormula("[" + field + "]-(" + value.ToString(CultureInfo.InvariantCulture) + ")", current[field] - value);
            }
        }
        var total = report.Rows["Total"];
        total["Date"] = context.BusinessDate;
        foreach (var field in BalanceFields.Concat(["SETTLE PNL"]).Concat(BalanceFields.Select(field => field + " DIFF")))
        {
            if (!report.Rows["TMA"].TryGetValue(field, out var tma) || !report.Rows["TMG"].TryGetValue(field, out var tmg)) continue;
            total[field] = new MonthlyFormula("[TMA:" + field + "]+[TMG:" + field + "]", Number(tma, field) + Number(tmg, field));
        }
        report.Rows["Break down"]["Date"] = context.BusinessDate;
        report.BalancesComplete = previous != null && previousPending != null && !report.Diagnostics.Any(diagnostic => diagnostic.Severity == "Incomplete");
        report.Provenance.Add(new
        {
            Kind = "Monthly field definitions", CurrentDate = context.BusinessDate,
            Bank = "Bank USD Equivalent; Account Type=Clients by Insto/Retail/Wholesale; Account Type=Settle without Settle Type restriction",
            Psp = "PSP USD Equivalent by source-derived Branch, including unmapped in-scope channels",
            Diff = "Complete current group minus complete previous group; disappearing/new keys included, not a sum over only current-row Movement",
            SettlePnlTma = TmaSettleTypes.ToArray(), SettlePnlTmg = new[] { "SETTLE" },
            SettlePnlFilter = "Organization=TMC, Branch, Type=Retail; no additional Source restriction",
        });
        return report;
    }

    static Dictionary<string, decimal> Aggregate(IEnumerable<Dictionary<string, object?>> bank,
        IEnumerable<Dictionary<string, object?>> psp, string branch)
    {
        var result = new Dictionary<string, decimal>(StringComparer.Ordinal)
        { ["Insto"] = 0m, ["Retail"] = 0m, ["Wholesale"] = 0m, ["PSP"] = 0m, ["Settle"] = 0m };
        foreach (var row in bank.Where(row => Text(row, "Branch") == branch))
        {
            var type = Text(row, "Account Type");
            var field = type == "SETTLE" ? "Settle" : type == "CLIENTS" ? Text(row, "Client Type") switch
            { "INSTO" => "Insto", "RETAIL" => "Retail", "WHOLESALE" => "Wholesale", _ => "" } : "";
            if (field.Length == 0) throw new InvalidDataException("月度Bank分类无效，无法将余额归入客户或结算账户");
            result[field] += Number(row.GetValueOrDefault("USD Equivalent"), "Bank USD Equivalent");
        }
        result["PSP"] = psp.Where(row => Text(row, "Branch") == branch)
            .Sum(row => Number(row.GetValueOrDefault("USD Equivalent"), "PSP USD Equivalent"));
        return result;
    }

    static bool IsSettleType(string branch, string type) => branch == "TMA" ? TmaSettleTypes.Contains(type) : type == "SETTLE";
    static string Organization(IReadOnlyDictionary<string, object?> row) =>
        Text(row, row.TryGetValue("Organization", out var value) && !string.IsNullOrWhiteSpace(value?.ToString()) ? "Organization" : "Company");
    static string Text(IReadOnlyDictionary<string, object?> row, string field) =>
        Regex.Replace(row.GetValueOrDefault(field)?.ToString()?.Trim() ?? "", @"\s+", " ").ToUpperInvariant();
    internal static decimal Number(object? value, string source) => value switch
    {
        decimal number => number, int number => number, double number => Convert.ToDecimal(number),
        FormulaCell formula => formula.CalculatedValue, MonthlyFormula formula => formula.CalculatedValue,
        _ => throw new InvalidDataException(source + " 缺少可靠的已计算数值"),
    };
}
