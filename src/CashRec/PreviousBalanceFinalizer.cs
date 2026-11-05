using System.Globalization;
using System.Text.Json;

namespace CashRec;

public sealed class PreviousBalanceFinalizer : IReportFinalizer
{
    static readonly string[] BalanceSheets = ["Equity", "Bank", "PSP"];
    sealed record Snapshot(DateTime Date, string Source, Dictionary<string, Dictionary<string, decimal>> Values);

    public void FinalizeReport(RunContext context, IList<ReportTable> tables, List<Diagnostic> diagnostics)
    {
        if (diagnostics.Any(d => d.Severity == "Error")) return;
        var balances = tables.Where(t => BalanceSheets.Contains(t.Name)).ToList();
        if (balances.Count != BalanceSheets.Length) throw new InvalidDataException("余额基础表不完整，无法匹配上期");
        foreach (var table in balances)
            if (table.Rows.Select(row => Key(table.Name, row)).Distinct().Count() != table.Rows.Count)
                throw new InvalidDataException(table.Name + "存在重复余额业务键，不能任选一条作为上期或当前余额");

        var previous = FindPrevious(context, diagnostics);
        var manual = ReadManualDep(context);
        var matchedManual = new HashSet<string>();
        if (previous == null)
            diagnostics.Add(new("Incomplete", "上期余额", "未找到此前有效的同口径余额基准；已保留当前余额，上期、Movement和PNL留空。请提供此前日期的原始余额输入或新版基础结果"));
        foreach (var table in balances)
        {
            EnsureColumns(table, "Prev USD Equivalent", "Movement", "Previous Date", "Previous Source");
            var missing = new List<string>();
            var index = 0;
            foreach (var row in table.Rows)
            {
                index++;
                var key = Key(table.Name, row);
                var current = Number(row.GetValueOrDefault("USD Equivalent"), table.Name + " 当前USD");
                decimal? dep = 0m;
                if (table.Name == "Equity")
                {
                    EnsureColumns(table, "DEP", "PNL");
                    if (manual.TryGetValue(key, out var saved)) { dep = saved; matchedManual.Add(key); }
                    row["DEP"] = dep;
                    row["PNL"] = null;
                }
                row["Prev USD Equivalent"] = row["Movement"] = row["Previous Date"] = row["Previous Source"] = null;
                if (previous == null) continue;
                if (!previous.Values[table.Name].TryGetValue(key, out var prior)) { missing.Add(key); continue; }
                var movement = current - prior;
                var excelRow = index + 1;
                row["Prev USD Equivalent"] = prior;
                row["Previous Date"] = previous.Date;
                row["Previous Source"] = previous.Source;
                row["Movement"] = new FormulaCell($"{Address(table, "USD Equivalent", excelRow)}-{Address(table, "Prev USD Equivalent", excelRow)}", movement);
                if (table.Name == "Equity")
                    row["PNL"] = new FormulaCell($"{Address(table, "Movement", excelRow)}-{Address(table, "DEP", excelRow)}", movement - (dep ?? 0m));
            }
            if (missing.Count > 0)
                diagnostics.Add(new("Incomplete", table.Name, $"上期{previous!.Date:yyyy.MM.dd}缺少{missing.Count}个业务键，不当零；例：{string.Join("；", missing.Take(5).Select(k => k.Replace('\u001f', '/')))}"));
        }
        var unmatched = manual.Where(pair => !matchedManual.Contains(pair.Key)).ToDictionary(pair => pair.Key, pair => pair.Value);
        if (unmatched.Count > 0)
        {
            var path = Path.Combine(context.DateDirectory, "Equity_UnmatchedDEP_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".json");
            File.WriteAllText(path, JsonSerializer.Serialize(unmatched, new JsonSerializerOptions { WriteIndented = true }));
            diagnostics.Add(new("Warning", "人工DEP", $"{unmatched.Count}个旧人工键本次不存在，已单独保留：{path}"));
        }
        if (previous != null)
            diagnostics.Add(new("Info", "上期余额", $"使用{previous.Date:yyyy.MM.dd}的同口径余额；实体重分类不作为当日资金变动"));
    }

    static Snapshot? FindPrevious(RunContext context, List<Diagnostic> diagnostics)
    {
        var parent = Directory.GetParent(context.DateDirectory)?.FullName;
        if (parent == null) return null;
        var candidates = Directory.EnumerateDirectories(parent)
            .Select(path => new { Path = path, Date = ParseFolderDate(Path.GetFileName(path)) })
            .Where(item => item.Date != null && item.Date < context.BusinessDate)
            .OrderByDescending(item => item.Date).ToList();
        foreach (var candidate in candidates)
        {
            var date = candidate.Date!.Value;
            var path = Path.Combine(candidate.Path, "CashRec_Base.xlsx");
            if (File.Exists(path))
            {
                try
                {
                    if (BalanceWorkbookReader.HasCurrentRules(path))
                    {
                        var values = new Dictionary<string, Dictionary<string, decimal>>();
                        foreach (var name in BalanceSheets)
                        {
                            var rows = BalanceWorkbookReader.ReadRows(path, name, SnapshotColumns(name));
                            if (rows.Any(row => row.GetValueOrDefault("Date") is not DateTime rowDate || rowDate.Date != date))
                                throw new InvalidDataException(name + "业务日期与上期目录不一致");
                            values[name] = Index(name, rows);
                        }
                        return new(date, path, values);
                    }
                    diagnostics.Add(new("Warning", "上期余额", $"{date:yyyy.MM.dd}已有结果不具备当前实体/估值口径标记，将尝试按当前规则重算原始余额"));
                }
                catch (Exception error) when (error is IOException or InvalidDataException or ArgumentException)
                { diagnostics.Add(new("Warning", "上期余额", $"无法使用{date:yyyy.MM.dd}已存余额：{error.Message}；尝试原始输入")); }
            }
            try
            {
                var prepared = WorkspaceInputs.Inspect(parent, date);
                string[] required = ["Equity", "BankAU", "BankGL", "PSP", "Rates"];
                if (required.Any(key => !File.Exists(prepared.Inputs[key])))
                    throw new InvalidDataException("缺少重算所需的权益/银行/PSP/汇率输入");
                var priorContext = new RunContext(date, candidate.Path, prepared.Inputs, context.MappingPath,
                    DailyRates.Load(prepared.Inputs["Rates"], date), context.Progress);
                var values = new Dictionary<string, Dictionary<string, decimal>>();
                foreach (var adapter in ReportRunner.Adapters().Where(adapter => BalanceSheets.Contains(adapter.Name)))
                {
                    context.Progress?.Invoke($"按当前规则重算上期{date:yyyy.MM.dd} {adapter.Name}");
                    var table = adapter.Process(priorContext);
                    var error = table.Diagnostics.FirstOrDefault(d => d.Severity == "Error");
                    if (error != null) throw new InvalidDataException(error.Source + "：" + error.Message);
                    values.Add(table.Name, Index(table.Name, table.Rows));
                }
                if (BalanceSheets.Any(name => !values.ContainsKey(name))) throw new InvalidDataException("重算余额模块不完整");
                diagnostics.Add(new("Info", "上期余额", $"{date:yyyy.MM.dd}原始余额已按当前监管、来源归属及币种规则重算；不依赖旧Branch或Group划分"));
                return new(date, candidate.Path + "（按当前规则重算）", values);
            }
            catch (Exception error) when (error is IOException or InvalidDataException or ArgumentException or InvalidOperationException)
            { diagnostics.Add(new("Warning", "上期余额", $"跳过无效基准{date:yyyy.MM.dd}：{error.Message}")); }
        }
        return null;
    }

    static Dictionary<string, decimal> Index(string table, IEnumerable<Dictionary<string, object?>> rows)
    {
        var result = new Dictionary<string, decimal>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var key = Key(table, row);
            if (!result.TryAdd(key, Number(row.GetValueOrDefault("USD Equivalent"), table + " 上期USD")))
                throw new InvalidDataException(table + "上期业务键重复：" + key.Replace('\u001f', '/'));
        }
        return result;
    }

    static Dictionary<string, decimal?> ReadManualDep(RunContext context)
    {
        var path = Path.Combine(context.DateDirectory, "CashRec_Base.xlsx");
        var result = new Dictionary<string, decimal?>(StringComparer.Ordinal);
        if (!File.Exists(path) || !BalanceWorkbookReader.HasCompatibleManualKeys(path) || !BalanceWorkbookReader.HasInitializedDep(path)) return result;
        foreach (var row in BalanceWorkbookReader.ReadRows(path, "Equity", ["Date", "Branch", "Retail Type", "Currency", "DEP"]))
        {
            if (row.GetValueOrDefault("Date") is not DateTime date || date.Date != context.BusinessDate)
                throw new InvalidDataException("现有Equity的日期与目录不一致，无法安全保留DEP");
            var key = Key("Equity", row);
            var value = row.GetValueOrDefault("DEP");
            if (!result.TryAdd(key, value == null ? null : Number(value, "现有人工DEP")))
                throw new InvalidDataException("现有Equity人工DEP业务键重复，不能任选一条");
        }
        return result;
    }

    static string Key(string table, IReadOnlyDictionary<string, object?> row)
    {
        string Text(string field)
        {
            var text = row.GetValueOrDefault(field)?.ToString()?.Trim() ?? "";
            if (text.Length == 0) throw new InvalidDataException(table + "余额缺少业务键字段 " + field);
            return text.ToUpperInvariant();
        }
        var branch = Text("Branch");
        if (branch is not ("TMA" or "TMG")) throw new InvalidDataException(table + "余额实体超出当前范围");
        if (table is "Bank" or "PSP")
        {
            var source = DataValue.Normalize(Text("Source"));
            var expectedBranch = source switch { "nabau" or "pspau" => "TMA", "nabgl" or "pspgl" => "TMG", _ => "" };
            if (expectedBranch != branch) throw new InvalidDataException(table + "余额Source与实体不符合当前规则");
        }
        return table switch
        {
            "Equity" => string.Join('\u001f', Text("Branch"), Text("Retail Type"), Text("Currency")),
            "Bank" => string.Join('\u001f', Text("Source"), DataValue.Account(Text("Account Number")), Text("Currency")),
            "PSP" => string.Join('\u001f', Text("Source"), Text("PSP"), Text("Currency")),
            _ => throw new InvalidDataException("不支持的余额表 " + table)
        };
    }
    static decimal Number(object? value, string label) => value switch
    {
        decimal number => number,
        FormulaCell formula => formula.CalculatedValue,
        double number => Convert.ToDecimal(number),
        int number => number,
        _ when value != null => DataValue.Decimal(value.ToString()!, label),
        _ => throw new InvalidDataException(label + "缺失，不当零")
    };
    static DateTime? ParseFolderDate(string text) => DateTime.TryParseExact(text, "yyyy.MM.dd", CultureInfo.InvariantCulture,
        DateTimeStyles.None, out var date) ? date.Date : null;
    static string[] SnapshotColumns(string table) => table switch
    {
        "Equity" => ["Date", "Branch", "Retail Type", "Currency", "USD Equivalent"],
        "Bank" => ["Date", "Branch", "Source", "Account Number", "Currency", "USD Equivalent"],
        "PSP" => ["Date", "Branch", "Source", "PSP", "Currency", "USD Equivalent"],
        _ => throw new InvalidDataException("不支持的余额表 " + table)
    };
    static void EnsureColumns(ReportTable table, params string[] fields)
    { foreach (var field in fields) if (!table.Columns.Contains(field)) table.Columns.Add(field); }
    static string Address(ReportTable table, string field, int row)
    {
        var column = table.Columns.IndexOf(field) + 1;
        if (column == 0) throw new InvalidDataException("缺少公式字段 " + field);
        var letters = "";
        while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
        return letters + row.ToString(CultureInfo.InvariantCulture);
    }
}
