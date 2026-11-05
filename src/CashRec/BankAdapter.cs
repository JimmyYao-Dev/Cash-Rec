using System.Globalization;
using ClosedXML.Excel;

namespace CashRec;

public sealed class BankAdapter : IReportAdapter
{
    public string Name => "Bank";

    readonly record struct MappingKey(string Source, string Account);
    sealed record AccountMapping(int Row, string AccountName, string Company, string Branch,
        string AccountType, string ClientType, string SettleType, string Bank)
    {
        public string Classification => string.Join("|", new[] { Company, AccountType, ClientType, SettleType }.Select(Canonical));
    }

    public ReportTable Process(RunContext context)
    {
        var table = new ReportTable(Name);
        table.Columns.AddRange(["Company", "Branch", "Source", "Account Name", "Account Number", "Currency",
            "EOD Balance", "Date", "Account Type", "Client Type", "Settle Type", "Bank", "Units per USD",
            "USD Equivalent", "Prev USD Equivalent", "Movement", "Source File", "Source Record", "Mapping Rows"]);
        var mappings = ReadMappings(context.MappingPath);
        var reported = new HashSet<MappingKey>();
        foreach (var (input, source, branch) in new[] { ("BankAU", "NAB AU", "TMA"), ("BankGL", "NAB GL", "TMG") })
        {
            var path = context.SourcePath(input);
            var records = 0;
            var skipped = 0;
            foreach (var row in CsvData.Read(path, "Account Number", ["Account Name", "Account Number", "Currency", "EOD balance"]))
            {
                records++;
                var account = DataValue.Account(DataValue.Get(row, "Account Number"));
                var accountName = DataValue.Get(row, "Account Name");
                if (account.Length == 0 || DataValue.Normalize(account) is "total" or "totalbalance") { skipped++; continue; }
                var currency = DataValue.Get(row, "Currency").ToUpperInvariant();
                var key = new MappingKey(CanonicalSource(source), account.ToUpperInvariant());
                AccountMapping? mapping = null;
                List<AccountMapping>? candidates = null;
                if (!mappings.TryGetValue(key, out candidates))
                {
                    table.Diagnostics.Add(new("Error", source, $"{Path.GetFileName(path)} 数据记录 {records}：账户 {account} 缺少映射，请补充 {source} + 账户号码"));
                }
                else if (candidates.Select(m => m.Classification).Distinct(StringComparer.Ordinal).Count() != 1)
                {
                    if (reported.Add(key)) table.Diagnostics.Add(new("Error", source,
                        $"账户 {account} 的业务分类冲突：BANK MAPPING 行 {string.Join(", ", candidates.Select(m => m.Row))}；未任选或重复应用分类"));
                }
                else
                {
                    var invalid = candidates.SelectMany(candidate => ClassificationErrors(candidate).Select(error => $"行 {candidate.Row}：{error}")).ToArray();
                    if (invalid.Length > 0)
                    {
                        if (reported.Add(key)) table.Diagnostics.Add(new("Error", source,
                            $"账户 {account} 的 BANK MAPPING 分类无效：{string.Join("；", invalid)}"));
                    }
                    else mapping = candidates[0];
                    if (mapping != null && reported.Add(key))
                    {
                        if (candidates.Count > 1) table.Diagnostics.Add(new("Info", source,
                            $"账户 {account} 的同类映射仅应用一次：BANK MAPPING 行 {string.Join(", ", candidates.Select(m => m.Row))}"));
                        var oldBranches = candidates.Select(m => m.Branch).Where(b => b.Length > 0 && !b.Equals(branch, StringComparison.OrdinalIgnoreCase)).Distinct().ToArray();
                        if (oldBranches.Length > 0) table.Diagnostics.Add(new("Warning", source,
                            $"账户 {account} 的旧映射 Branch={string.Join("/", oldBranches)}；按输入来源归 {branch}，映射行 {string.Join(", ", candidates.Select(m => m.Row))}"));
                    }
                }
                decimal? amount = null, rate = null, usd = null;
                try
                {
                    if (currency.Length == 0) throw new InvalidDataException("Currency 为空");
                    amount = DataValue.Decimal(DataValue.Get(row, "EOD balance"), source + " EOD balance");
                    rate = context.Rates.Get(currency);
                    usd = amount.Value / rate.Value;
                }
                catch (InvalidDataException error)
                { table.Diagnostics.Add(new("Error", source, $"数据记录 {records} 账户 {account}：{error.Message}")); }
                table.AddRow(new()
                {
                    ["Company"] = mapping?.Company,
                    ["Branch"] = branch,
                    ["Source"] = source,
                    ["Account Name"] = accountName,
                    ["Account Number"] = account,
                    ["Currency"] = currency,
                    ["EOD Balance"] = amount,
                    ["Date"] = context.BusinessDate,
                    ["Account Type"] = mapping?.AccountType,
                    ["Client Type"] = mapping?.ClientType,
                    ["Settle Type"] = mapping?.SettleType,
                    ["Bank"] = mapping?.Bank,
                    ["Units per USD"] = rate,
                    ["USD Equivalent"] = usd,
                    ["Prev USD Equivalent"] = null,
                    ["Movement"] = null,
                    ["Source File"] = path,
                    ["Source Record"] = records,
                    ["Mapping Rows"] = candidates == null ? "" : string.Join(",", candidates.Select(m => m.Row))
                });
            }
            if (records == 0) table.Diagnostics.Add(new("Warning", source, "CSV表头有效，但没有银行账户数据记录"));
            table.Diagnostics.Add(new("Info", source, $"读取 {records} 条表头后记录，排除 {skipped} 条总计或非账户记录；账户余额不按相同金额去重"));
        }
        return table;
    }

    static Dictionary<MappingKey, List<AccountMapping>> ReadMappings(string path)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheets.FirstOrDefault(s => DataValue.Normalize(s.Name) == "bankmapping")
            ?? throw new InvalidDataException("映射：缺少 BANK MAPPING 工作表");
        var header = sheet.RowsUsed().Take(30).FirstOrDefault(row => row.CellsUsed().Any(c => DataValue.Normalize(c.GetString()) == "accountnumber")
            && row.CellsUsed().Any(c => DataValue.Normalize(c.GetString()) == "sourcefile"))
            ?? throw new InvalidDataException("BANK MAPPING：未找到 Account Number 与 Source File 表头");
        var columns = header.CellsUsed().ToDictionary(c => DataValue.Normalize(c.GetString()), c => c.Address.ColumnNumber);
        foreach (var required in new[] { "accountnumber", "sourcefile", "company", "accounttype", "clienttype", "settletype" })
            if (!columns.ContainsKey(required)) throw new InvalidDataException("BANK MAPPING 缺少字段：" + required);
        var result = new Dictionary<MappingKey, List<AccountMapping>>();
        foreach (var row in sheet.RowsUsed().Where(r => r.RowNumber() > header.RowNumber()))
        {
            string Text(string key) => columns.TryGetValue(key, out var column) ? row.Cell(column).GetString().Trim() : "";
            var source = CanonicalSource(Text("sourcefile"));
            if (source is not ("nabau" or "nabgl")) continue;
            var accountCell = row.Cell(columns["accountnumber"]);
            var account = DataValue.Account(accountCell.DataType == XLDataType.Number
                ? accountCell.GetFormattedString(CultureInfo.InvariantCulture) : accountCell.GetString()).ToUpperInvariant();
            if (account.Length == 0) continue;
            var key = new MappingKey(source, account);
            if (!result.TryGetValue(key, out var list)) result[key] = list = [];
            list.Add(new(row.RowNumber(), Text("accountname"), Text("company"), Text("branch"), Text("accounttype"), Text("clienttype"), Text("settletype"), Text("bank")));
        }
        return result;
    }

    static string Canonical(string value) => string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToUpperInvariant();
    static string CanonicalSource(string value) => DataValue.Normalize(Path.GetFileNameWithoutExtension(value.Replace('\\', '/').Trim()));
    static IEnumerable<string> ClassificationErrors(AccountMapping mapping)
    {
        if (string.IsNullOrWhiteSpace(mapping.Company)) yield return "Company不能为空";
        var type = Canonical(mapping.AccountType);
        if (type is not ("CLIENTS" or "SETTLE")) yield return "Account Type必须为Clients或Settle";
        if (type == "CLIENTS" && Canonical(mapping.ClientType) is not ("INSTO" or "RETAIL" or "WHOLESALE"))
            yield return "Clients账户的Client Type必须为Insto、Retail或Wholesale";
        if (type == "SETTLE" && string.IsNullOrWhiteSpace(mapping.SettleType))
            yield return "Settle账户的Settle Type不能为空（合法NA保留）";
    }
}
