using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace CashRec;

/// <summary>Prepared BANK BD transactions, kept distinct from MT4 client flows.</summary>
public sealed class BankBdTransactionAdapter : IReportAdapter
{
    public string Name => "TransactionExchange";

    static readonly string[] RequiredHeaders = ["Comment", "Rational", "Currency", "Company", "Amount"];
    static readonly Dictionary<string, string> BusinessTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Retail"] = "Retail", ["Wholesale"] = "Wholesale", ["Insto"] = "Insto",
        ["PSP"] = "PSP", ["Settle"] = "Settle", ["Bank Fee"] = "Bank Fee",
        ["Transfer"] = "Transfer", ["Other"] = "Other",
    };
    static readonly Dictionary<string, string> TransactionTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Deposit"] = "Deposit", ["Dep"] = "Deposit",
        ["Withdrawal"] = "Withdrawal", ["Withdraw"] = "Withdrawal",
        ["Adjustment"] = "Adjustment", ["Settle"] = "Settle",
        ["Settlement"] = "Settle", ["Bank Fee"] = "Bank Fee",
        ["Transfer"] = "Transfer", ["Other"] = "Other", ["Others"] = "Other",
    };

    public ReportTable Process(RunContext context)
    {
        var path = context.SourcePath("BankBD");
        var table = new ReportTable(Name);
        table.Columns.AddRange([
            "Source", "Comment", "Rational", "Company", "Branch", "Type",
            "Transaction Type", "Exchange Type", "Date", "Amount", "Currency",
            "Units per USD", "USD Equivalent", "Source File", "Source Sheet",
            "Source Row", "Source Date", "Classification Source",
        ]);
        try
        {
            using var book = new XLWorkbook(path);
            var input = FindInputTable(book);
            var excluded = new Dictionary<string, int>(StringComparer.Ordinal);
            var otherTypes = new Dictionary<string, int>(StringComparer.Ordinal);
            var rowCount = 0;
            var templateCount = 0;
            foreach (var row in input.Sheet.RowsUsed().Where(r => r.RowNumber() > input.HeaderRow))
            {
                var source = $"{Path.GetFileName(path)} / {input.Sheet.Name} 第 {row.RowNumber()} 行";
                try
                {
                    string Text(string header) => ReadText(row.Cell(input.Headers[DataValue.Normalize(header)]), source);
                    var comment = Text("Comment");
                    var rational = Text("Rational");
                    var currency = Text("Currency").Trim().ToUpperInvariant();
                    if (string.IsNullOrWhiteSpace(comment) && string.IsNullOrWhiteSpace(rational) && currency.Length == 0
                        && IsBlankOrZero(row.Cell(input.Headers["amount"]), source + " Amount")
                        && new[] { "credit", "debit" }.All(name => !input.Headers.TryGetValue(name, out var column)
                            || IsBlankOrZero(row.Cell(column), source + " " + name)))
                    {
                        templateCount++;
                        continue;
                    }
                    rowCount++;
                    if (rowCount % 10000 == 0)
                        context.Progress?.Invoke($"处理 BANK BD：{rowCount:N0} 条源记录");
                    var company = Text("Company");
                    var match = Regex.Match(company.Trim(), @"^(TMA|TMG)(?:\s*-\s*(.*))?$",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (!match.Success)
                    {
                        var reason = string.IsNullOrWhiteSpace(company) ? "Company 实体为空" : "Company 实体不在 TMA/TMG 范围";
                        excluded[reason] = excluded.GetValueOrDefault(reason) + 1;
                        continue;
                    }
                    var branch = match.Groups[1].Value.ToUpperInvariant();
                    var originalType = CollapseSpaces(match.Groups[2].Value);
                    var type = BusinessTypes.GetValueOrDefault(originalType, "Other");
                    if (!BusinessTypes.ContainsKey(originalType))
                        otherTypes[originalType.Length == 0 ? "(空)" : originalType] =
                            otherTypes.GetValueOrDefault(originalType.Length == 0 ? "(空)" : originalType) + 1;
                    if (currency.Length == 0) throw new InvalidDataException($"{source}：Currency 为空");
                    var amount = ExcelNumbers.Read(row.Cell(input.Headers["amount"]), source + " Amount");
                    var rate = context.Rates.Get(currency);
                    var transactionType = ClassifyTransaction(rational, comment, type);
                    var sourceDate = input.Headers.TryGetValue("date", out var dateColumn)
                        ? ReadText(row.Cell(dateColumn), source) : "";
                    table.AddRow(new Dictionary<string, object?>
                    {
                        ["Source"] = "BANK BD",
                        ["Comment"] = comment,
                        ["Rational"] = rational,
                        ["Company"] = company,
                        ["Branch"] = branch,
                        ["Type"] = type,
                        ["Transaction Type"] = transactionType,
                        ["Exchange Type"] = transactionType is "PSP CN Commission" or "PSP Commission"
                            ? transactionType : "BANK BD",
                        ["Date"] = context.BusinessDate,
                        ["Amount"] = amount,
                        ["Currency"] = currency,
                        ["Units per USD"] = rate,
                        ["USD Equivalent"] = amount / rate,
                        ["Source File"] = Path.GetFileName(path),
                        ["Source Sheet"] = input.Sheet.Name,
                        ["Source Row"] = row.RowNumber(),
                        ["Source Date"] = sourceDate,
                        ["Classification Source"] = "Company + Rational / Comment",
                    });
                }
                catch (Exception error) when (error is InvalidDataException or InvalidCastException or OverflowException)
                {
                    table.Diagnostics.Add(new("Error", "BANK BD", $"{source}：{error.Message}"));
                }
            }
            foreach (var (reason, count) in excluded)
                table.Diagnostics.Add(new("Warning", "BANK BD", $"排除 {count} 条记录：{reason}"));
            foreach (var (original, count) in otherTypes)
                table.Diagnostics.Add(new("Warning", "BANK BD", $"{count} 条记录的业务类别 {original} 归入 Other，原值保留"));
            if (templateCount > 0)
                table.Diagnostics.Add(new("Info", "BANK BD", $"跳过 {templateCount} 条无业务内容的空模板行"));
            table.Diagnostics.Add(new("Info", "BANK BD",
                $"读取 {rowCount} 条准备记录，保留 {table.Rows.Count} 条 TMA/TMG 交易；使用源 Amount 符号"));
            return table;
        }
        catch (Exception error)
        {
            throw new InvalidDataException($"BANK BD / {Path.GetFileName(path)}：{error.Message}", error);
        }
    }

    static string CollapseSpaces(string value) => Regex.Replace(value.Trim(), @"\s+", " ");

    static bool IsBlankOrZero(IXLCell cell, string source) =>
        (!cell.HasFormula && cell.Value.Type == XLDataType.Blank) || ExcelNumbers.Read(cell, source) == 0m;

    static string ClassifyTransaction(string rational, string comment, string type)
    {
        foreach (var value in new[] { CollapseSpaces(rational), CollapseSpaces(comment) })
        {
            if (value.Equals("PSP CN Commission", StringComparison.OrdinalIgnoreCase)) return "PSP CN Commission";
            if (value.Equals("PSP Commission", StringComparison.OrdinalIgnoreCase)) return "PSP Commission";
        }
        var original = CollapseSpaces(rational);
        var category = TransactionTypes.GetValueOrDefault(original, "Other");
        return (type, category) switch
        {
            ("PSP", "Deposit") => "PSP Dep",
            ("PSP", "Withdrawal") => "PSP WITH",
            _ => category,
        };
    }


    static string ReadText(IXLCell cell, string source)
    {
        if (!cell.HasFormula) return cell.GetString();
        if (cell.CachedValue.Type is XLDataType.Blank or XLDataType.Error)
            throw new InvalidDataException($"{source} {cell.Address} 公式缺少有效缓存");
        return cell.CachedValue.ToString(CultureInfo.InvariantCulture);
    }

    sealed record InputTable(IXLWorksheet Sheet, int HeaderRow, Dictionary<string, int> Headers);

    static InputTable FindInputTable(XLWorkbook book)
    {
        var candidates = new List<InputTable>();
        foreach (var sheet in book.Worksheets)
        foreach (var row in sheet.RowsUsed().Take(30))
        {
            var headers = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var cell in row.CellsUsed())
            {
                if (cell.HasFormula) continue;
                var name = DataValue.Normalize(cell.GetString());
                if (name.Length == 0) continue;
                if (!headers.TryAdd(name, cell.Address.ColumnNumber))
                    headers[name] = -1;
            }
            if (!RequiredHeaders.All(name => headers.TryGetValue(DataValue.Normalize(name), out var column) && column > 0))
                continue;
            candidates.Add(new(sheet, row.RowNumber(), headers));
            break;
        }
        if (candidates.Count == 0)
            throw new InvalidDataException("未找到唯一的 Comment、Rational、Currency、Company、Amount 交易表头");
        if (candidates.Count == 1) return candidates[0];
        var named = candidates.Where(c => DataValue.Normalize(c.Sheet.Name) == "bankbd").ToList();
        if (named.Count == 1) return named[0];
        throw new InvalidDataException("存在多个 BANK BD 交易明细表，无法确定输入");
    }
}

