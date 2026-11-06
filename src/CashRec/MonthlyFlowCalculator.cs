using System.Text.RegularExpressions;

namespace CashRec;

/// <summary>Signed prepared-window flows, separated from balance snapshots and sources.</summary>
public static class MonthlyFlowCalculator
{
    public static readonly string[] Fields = ["PSP DEP", "PSP WITH", "NET PSP", "NET BW", "NET WS"];

    public static void Apply(MonthlyReport report, IEnumerable<ReportTable> tables, Action<string>? progress = null)
    {
        var sums = new[] { "TMA", "TMG" }.ToDictionary(branch => branch,
            _ => Fields.ToDictionary(field => field, _ => 0m));
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var processed = 0;
        foreach (var row in tables.Where(table => table.Name == "TransactionExchange").SelectMany(table => table.Rows))
        {
            if (++processed % 10000 == 0) progress?.Invoke($"汇总客户净额：{processed:N0} 条准备记录");
            var branch = Text(row, "Branch");
            if (!sums.ContainsKey(branch)) continue;
            if (row.GetValueOrDefault("Date") is not DateTime date)
                throw new InvalidDataException("月度客户净额：交易缺少有效报表归属日期");
            if (date.Date != report.Date) continue;
            var source = Text(row, "Source");
            var type = Text(row, "Type");
            var transactionType = Text(row, "Transaction Type");
            string? field = null;
            if (source == "TRANSACTION REPORTS" && type == "PSP")
                field = transactionType switch { "PSP DEP" => "PSP DEP", "PSP WITH" => "PSP WITH", _ => null };
            else if (source == "TRANSACTION REPORTS" && type == "RETAIL" && Text(row, "Exchange Type") == "MT4 BW")
                field = "NET BW";
            else if (source == "BANK BD" && type == "WHOLESALE") field = "NET WS";
            if (field == null) continue;
            sums[branch][field] += MonthlyBalanceCalculator.Number(row.GetValueOrDefault("USD Equivalent"), "月度 " + field + " USD Equivalent");
            var key = branch + "/" + field;
            counts[key] = counts.GetValueOrDefault(key) + 1;
        }
        foreach (var branch in new[] { "TMA", "TMG" })
        {
            var values = sums[branch];
            values["NET PSP"] = values["PSP DEP"] + values["PSP WITH"];
            foreach (var field in Fields) report.Rows[branch][field] = field == "NET PSP"
                ? new MonthlyFormula("[PSP DEP]+[PSP WITH]", values[field]) : values[field];
        }
        foreach (var field in Fields)
            report.Rows["Total"][field] = new MonthlyFormula("[TMA:" + field + "]+[TMG:" + field + "]", sums["TMA"][field] + sums["TMG"][field]);
        report.FlowsComplete = true;
        report.Provenance.Add(new
        {
            Kind = "Monthly client flows", Date = report.Date, MatchedRecords = counts,
            Psp = "Source=Transaction reports, Type=PSP, Transaction Type=PSP Dep or PSP WITH",
            NetPsp = "Signed PSP DEP + PSP WITH; no sign reversal",
            BankWire = "Source=Transaction reports, Type=Retail, Exchange Type=MT4 BW",
            Wholesale = "Source=BANK BD, Type=Wholesale",
            Amount = "Prepared report Date, Branch and USD Equivalent; all distinct records retained, including equal amounts and zeros",
        });
    }

    static string Text(IReadOnlyDictionary<string, object?> row, string field) =>
        Regex.Replace(row.GetValueOrDefault(field)?.ToString()?.Trim() ?? "", @"\s+", " ").ToUpperInvariant();
}
