namespace CashRec;

/// <summary>Trading-side Deposit/Withdrawal records from prepared MT4 exports.</summary>
public sealed class Mt4TransactionAdapter : IReportAdapter
{
    const string Source = "Transaction reports";
    public string Name => "TransactionExchange";

    public ReportTable Process(RunContext context)
    {
        var path = context.SourcePath("Transactions");
        var table = new ReportTable(Name);
        table.Columns.AddRange([
            "Comment", "Rational", "Currency", "Company", "Amount", "Branch", "Type",
            "Units per USD", "Transaction Type", "Exchange Type", "Source", "MT4_ID",
            "USD Equivalent", "Date", "Source Date", "Source File", "Source Record",
            "Transaction ID", "Server Transaction ID", "User ID", "Trading Server",
            "Payment Amount", "Payment Currency", "Accounting Comment",
        ]);
        var excluded = new Dictionary<string, int>(StringComparer.Ordinal);
        var sourceRecords = 0;
        var headerChecked = false;

        foreach (var row in CsvData.Read(path))
        {
            sourceRecords++;
            if (!headerChecked)
            {
                RequireHeaders(row, path);
                headerChecked = true;
            }
            if (sourceRecords % 10000 == 0)
                context.Progress?.Invoke($"处理 MT4 交易：{sourceRecords:N0} 条源记录");

            var branch = EntityRules.RegulatorBranch(DataValue.Get(row, "jurisdiction"), out var reason);
            if (branch == null)
            {
                CountExclusion(reason);
                continue;
            }
            var accountingComment = Original(row, "accounting_comment");
            var deposit = accountingComment.Trim().Equals("Deposit", StringComparison.OrdinalIgnoreCase);
            var withdrawal = accountingComment.Trim().Equals("Withdrawal", StringComparison.OrdinalIgnoreCase);
            if (!deposit && !withdrawal)
            {
                CountExclusion("accounting_comment 不是 Deposit/Withdrawal");
                continue;
            }

            var currency = DataValue.Get(row, "trading_server_currency").ToUpperInvariant();
            if (currency.Length == 0)
                throw new InvalidDataException($"{Path.GetFileName(path)} 第 {sourceRecords} 条记录：trading_server_currency 为空");
            var amount = DataValue.Decimal(DataValue.Get(row, "fund_amount"),
                $"{Path.GetFileName(path)} 第 {sourceRecords} 条记录 fund_amount");
            decimal rate;
            try { rate = context.Rates.Get(currency); }
            catch (InvalidDataException error)
            {
                throw new InvalidDataException($"{Path.GetFileName(path)} 第 {sourceRecords} 条记录：{error.Message}", error);
            }
            var comment = Original(row, "comment");
            var bankWire = comment.Contains("NAB", StringComparison.OrdinalIgnoreCase);
            table.AddRow(new()
            {
                ["Comment"] = comment,
                ["Rational"] = accountingComment,
                ["Currency"] = currency,
                ["Company"] = "TMC",
                ["Amount"] = amount,
                ["Branch"] = branch,
                ["Type"] = bankWire ? "Retail" : "PSP",
                ["Units per USD"] = rate,
                ["Transaction Type"] = bankWire
                    ? deposit ? "BW Dep" : "BW WD"
                    : deposit ? "PSP Dep" : "PSP WITH",
                ["Exchange Type"] = bankWire ? "MT4 BW" : "",
                ["Source"] = Source,
                ["MT4_ID"] = Original(row, "trading_server_login"),
                ["USD Equivalent"] = amount / rate,
                ["Date"] = context.BusinessDate,
                ["Source Date"] = Original(row, "date"),
                ["Source File"] = path,
                ["Source Record"] = sourceRecords,
                ["Transaction ID"] = Original(row, "crm_transaction_no"),
                ["Server Transaction ID"] = Original(row, "trading_server_transaction_no"),
                ["User ID"] = Original(row, "user_id"),
                ["Trading Server"] = Original(row, "trading_server"),
                ["Payment Amount"] = Original(row, "payment_amount"),
                ["Payment Currency"] = Original(row, "payment_currency"),
                ["Accounting Comment"] = accountingComment,
            });
        }

        foreach (var (reason, count) in excluded.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            table.Diagnostics.Add(new("Warning", Source, $"排除 {count} 条记录：{reason}"));
        table.Diagnostics.Add(new("Info", Source,
            $"读取 {sourceRecords} 条准备记录，保留 {table.Rows.Count} 条 MT4 入出金；金额保持源符号，源日期不作二次筛选"));
        return table;

        void CountExclusion(string reason) => excluded[reason] = excluded.GetValueOrDefault(reason) + 1;
    }

    static string Original(IReadOnlyDictionary<string, string> row, string name)
    {
        if (row.TryGetValue(name, out var value)) return value;
        var key = row.Keys.FirstOrDefault(key => DataValue.Normalize(key) == DataValue.Normalize(name));
        return key == null ? "" : row[key];
    }

    static void RequireHeaders(IReadOnlyDictionary<string, string> row, string path)
    {
        string[] required = ["date", "jurisdiction", "accounting_comment", "comment", "fund_amount", "trading_server_currency"];
        var missing = required.Where(name => !row.Keys.Any(key =>
            DataValue.Normalize(key) == DataValue.Normalize(name))).ToArray();
        if (missing.Length > 0)
            throw new InvalidDataException($"{Path.GetFileName(path)}：缺少 MT4 交易表头 {string.Join(", ", missing)}");
    }
}
