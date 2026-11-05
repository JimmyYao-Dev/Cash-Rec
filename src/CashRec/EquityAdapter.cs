namespace CashRec;

/// <summary>Prepared-file Equity balances; original report dates are provenance.</summary>
public sealed class EquityAdapter : IReportAdapter
{
    public string Name => "Equity";

    readonly record struct EquityKey(string Branch, string RetailType, string Currency);

    sealed class Balance
    {
        public decimal Amount;
        public int Records;
        public SortedSet<string> SourceDates { get; } = new(StringComparer.Ordinal);
    }

    public ReportTable Process(RunContext context)
    {
        var path = context.SourcePath("Equity");
        var table = new ReportTable(Name);
        table.Columns.AddRange([
            "Company", "Branch", "Client Type", "Retail Type", "Trading Type", "Currency",
            "Amount", "DEP", "PNL", "Date", "Units per USD", "USD Equivalent",
            "Prev USD Equivalent", "Movement", "Trend", "Source File", "Source Row Count",
            "Source Dates", "Source Filter",
        ]);
        var groups = new Dictionary<EquityKey, Balance>();
        var excluded = new Dictionary<string, int>(StringComparer.Ordinal);
        var sourceRecords = 0;
        var headerChecked = false;

        foreach (var source in CsvData.Read(path))
        {
            sourceRecords++;
            if (!headerChecked)
            {
                RequireHeaders(source, path);
                headerChecked = true;
            }
            if (sourceRecords % 10000 == 0)
                context.Progress?.Invoke($"处理 Equity：{sourceRecords:N0} 条源记录");

            var branch = EntityRules.RegulatorBranch(DataValue.Get(source, "report_jurisdiction"), out var reason);
            if (branch == null)
            {
                excluded[reason] = excluded.GetValueOrDefault(reason) + 1;
                continue;
            }
            if (!bool.TryParse(DataValue.Get(source, "is_ib"), out var isIb))
            {
                table.Diagnostics.Add(new("Error", Name,
                    $"{Path.GetFileName(path)} 第 {sourceRecords} 条记录：is_ib 必须为 true 或 false"));
                continue;
            }
            var currency = DataValue.Get(source, "currency").ToUpperInvariant();
            if (currency.Length == 0)
            {
                table.Diagnostics.Add(new("Error", Name,
                    $"{Path.GetFileName(path)} 第 {sourceRecords} 条记录：currency 为空"));
                continue;
            }
            decimal amount;
            try
            {
                amount = DataValue.Decimal(DataValue.Get(source, "positive_equity"),
                    $"{Path.GetFileName(path)} 第 {sourceRecords} 条记录 positive_equity");
            }
            catch (InvalidDataException error)
            {
                table.Diagnostics.Add(new("Error", Name, error.Message));
                continue;
            }

            var key = new EquityKey(branch, isIb ? "IB" : "Standard", currency);
            if (!groups.TryGetValue(key, out var group))
            {
                group = new Balance();
                groups.Add(key, group);
            }
            group.Amount += amount;
            group.Records++;
            var sourceDate = DataValue.Get(source, "dt_report");
            group.SourceDates.Add(sourceDate.Length == 0 ? "(空)" : sourceDate);
        }

        foreach (var (key, group) in groups.OrderBy(pair => pair.Key.Branch, StringComparer.Ordinal)
            .ThenBy(pair => pair.Key.RetailType, StringComparer.Ordinal)
            .ThenBy(pair => pair.Key.Currency, StringComparer.Ordinal))
        {
            var rate = context.Rates.Get(key.Currency);
            table.AddRow(new()
            {
                ["Company"] = "TMC",
                ["Branch"] = key.Branch,
                ["Client Type"] = "Retail",
                ["Retail Type"] = key.RetailType,
                ["Trading Type"] = "MT4",
                ["Currency"] = key.Currency,
                ["Amount"] = group.Amount,
                ["DEP"] = null,
                ["PNL"] = null,
                ["Date"] = context.BusinessDate,
                ["Units per USD"] = rate,
                ["USD Equivalent"] = group.Amount / rate,
                ["Prev USD Equivalent"] = null,
                ["Movement"] = null,
                ["Trend"] = null,
                ["Source File"] = path,
                ["Source Row Count"] = group.Records,
                ["Source Dates"] = string.Join("; ", group.SourceDates),
                ["Source Filter"] = $"report_jurisdiction => {key.Branch}; is_ib => {key.RetailType}; currency={key.Currency}",
            });
        }
        foreach (var (reason, count) in excluded.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            table.Diagnostics.Add(new("Warning", Name, $"排除 {count} 条记录：{reason}"));
        table.Diagnostics.Add(new("Info", Name,
            $"读取 {sourceRecords} 条准备记录，生成 {table.Rows.Count} 个实体/IB状态/币种余额组；源日期不作二次筛选"));
        return table;
    }

    static void RequireHeaders(IReadOnlyDictionary<string, string> row, string path)
    {
        string[] required = ["dt_report", "report_jurisdiction", "is_ib", "positive_equity", "currency"];
        var missing = required.Where(name => !row.Keys.Any(key =>
            DataValue.Normalize(key) == DataValue.Normalize(name))).ToArray();
        if (missing.Length > 0)
            throw new InvalidDataException($"{Path.GetFileName(path)}：缺少 Equity 表头 {string.Join(", ", missing)}");
    }
}
