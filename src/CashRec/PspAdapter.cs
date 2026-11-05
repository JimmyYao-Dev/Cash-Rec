using System.Globalization;
using ClosedXML.Excel;

namespace CashRec;

public sealed class PspAdapter : IReportAdapter
{
    public string Name => "PSP";
    sealed record Mapping(string Company, string Branch, int Row);
    sealed record Layout(IXLWorksheet Sheet, int Header, int Source, int Way, int Currency, int Amount);

    public ReportTable Process(RunContext context)
    {
        var table = new ReportTable(Name);
        table.Columns.AddRange(["Company", "Branch", "Source", "PSP", "Currency", "Amount", "Date",
            "Units per USD", "USD Equivalent", "Prev USD Equivalent", "Movement", "Mapping Status",
            "Source File", "Source Sheet", "Source Row", "Mapping Rows"]);
        var mappings = ReadMappings(context.MappingPath, table.Diagnostics);
        using var workbook = new XLWorkbook(context.SourcePath("PSP"));
        var layouts = new List<Layout>();
        foreach (var sheet in workbook.Worksheets.Where(s => MatchesMonth(s.Name, context.BusinessDate)))
        {
            foreach (var row in sheet.RowsUsed().Take(20))
            {
                var cells = row.CellsUsed().ToList();
                int Column(params string[] names) => cells.FirstOrDefault(c => names.Contains(DataValue.Normalize(c.GetString())))?.Address.ColumnNumber ?? 0;
                var source = Column("source");
                var way = Column("way", "psp");
                var currency = Column("curr", "currency");
                if (source == 0 || way == 0 || currency == 0) continue;
                foreach (var amount in cells.Where(c => DataValue.Normalize(c.GetString()) == "amount"))
                {
                    var column = amount.Address.ColumnNumber;
                    if (Enumerable.Range(1, row.RowNumber() - 1).Any(r => IsDate(sheet.Cell(r, column), context.BusinessDate)))
                        layouts.Add(new(sheet, row.RowNumber(), source, way, currency, column));
                }
                break;
            }
        }
        if (layouts.Count != 1)
            throw new InvalidDataException($"PSP Balance Summary：{context.BusinessDate:yyyy.MM.dd} 的月份/Amount列匹配数为 {layouts.Count}，需要唯一对应日期列，不能改用邻日或 Eq USD");
        var layout = layouts[0];
        var excluded = new Dictionary<string, int>();
        foreach (var row in layout.Sheet.RowsUsed().Where(r => r.RowNumber() > layout.Header))
        {
            var source = row.Cell(layout.Source).GetString().Trim().ToUpperInvariant();
            var way = row.Cell(layout.Way).GetString().Trim();
            if (source.Length == 0 && way.Length == 0) continue;
            if (IsTotal(source) || IsTotal(way)) { Count("Total汇总行"); continue; }
            var branch = source switch { "PSPAU" => "TMA", "PSPGL" => "TMG", _ => null };
            if (branch == null) { Count("非本版Source " + source); continue; }
            var currency = row.Cell(layout.Currency).GetString().Trim().ToUpperInvariant();
            if (way.Length == 0 || currency.Length == 0)
                throw new InvalidDataException($"PSP {layout.Sheet.Name}!{row.RowNumber()}：WAY/CURR不能为空");
            var amount = ReadAmount(row.Cell(layout.Amount), $"PSP {layout.Sheet.Name} 第{row.RowNumber()}行 Amount");
            var rate = context.Rates.Get(currency);
            mappings.TryGetValue(Key(source, way, currency), out var matches);
            var company = matches?.Select(m => m.Company).Distinct(StringComparer.OrdinalIgnoreCase).ToArray() ?? [];
            var status = matches == null ? "待补映射" : company.Length > 1 ? "映射属性冲突待核对" : "已匹配";
            if (matches == null)
                table.Diagnostics.Add(new("Warning", Name, $"新增PSP {source}/{way}/{currency}：金额保留，实体为{branch}，其他映射属性待补"));
            else
            {
                if (company.Length > 1)
                    table.Diagnostics.Add(new("Warning", Name, $"PSP MAPPING {source}/{way}/{currency} 的Company冲突，行{string.Join(",", matches.Select(m => m.Row))}；金额只计一次，Company留空待核对"));
                if (matches.Any(m => !string.Equals(m.Branch, branch, StringComparison.OrdinalIgnoreCase)))
                    table.Diagnostics.Add(new("Warning", Name, $"PSP MAPPING {source}/{way}/{currency} 的Branch与Source冲突，仍归{branch}；行{string.Join(",", matches.Select(m => m.Row))}"));
            }
            table.AddRow(new()
            {
                ["Company"] = company.Length == 1 ? company[0] : null,
                ["Branch"] = branch, ["Source"] = source, ["PSP"] = way, ["Currency"] = currency,
                ["Amount"] = amount, ["Date"] = context.BusinessDate, ["Units per USD"] = rate,
                ["USD Equivalent"] = amount / rate, ["Prev USD Equivalent"] = null, ["Movement"] = null,
                ["Mapping Status"] = status, ["Source File"] = context.SourcePath("PSP"),
                ["Source Sheet"] = layout.Sheet.Name, ["Source Row"] = row.RowNumber(),
                ["Mapping Rows"] = matches == null ? "" : string.Join(",", matches.Select(m => m.Row))
            });
        }
        foreach (var (reason, count) in excluded)
            table.Diagnostics.Add(new("Info", Name, $"排除{count}行：{reason}"));
        table.Diagnostics.Add(new("Info", Name, $"仅读取{context.BusinessDate:yyyy.MM.dd}的Amount，保留{table.Rows.Count}条PSP余额"));
        return table;

        void Count(string reason) => excluded[reason] = excluded.GetValueOrDefault(reason) + 1;
    }

    static Dictionary<string, List<Mapping>> ReadMappings(string path, List<Diagnostic> diagnostics)
    {
        var result = new Dictionary<string, List<Mapping>>(StringComparer.OrdinalIgnoreCase);
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheets.FirstOrDefault(s => s.Name.Equals("PSP MAPPING", StringComparison.OrdinalIgnoreCase));
        if (sheet == null) { diagnostics.Add(new("Warning", "PSP", "缺少PSP MAPPING；已知Source余额仍保留，映射属性待补")); return result; }
        var header = sheet.RowsUsed().FirstOrDefault() ?? throw new InvalidDataException("PSP MAPPING为空");
        var columns = header.CellsUsed().ToDictionary(c => DataValue.Normalize(c.GetString()), c => c.Address.ColumnNumber);
        foreach (var field in new[] { "source", "psp", "currency" })
            if (!columns.ContainsKey(field)) throw new InvalidDataException("PSP MAPPING缺少表头 " + field);
        foreach (var row in sheet.RowsUsed().Where(r => r.RowNumber() > header.RowNumber()))
        {
            string Value(string field) => columns.TryGetValue(field, out var col) ? row.Cell(col).GetString().Trim() : "";
            var source = Value("source"); var psp = Value("psp"); var currency = Value("currency");
            if (source.Length == 0 && psp.Length == 0) continue;
            var key = Key(source, psp, currency);
            if (!result.TryGetValue(key, out var entries)) result[key] = entries = [];
            entries.Add(new(Value("company"), Value("branch"), row.RowNumber()));
        }
        return result;
    }

    static string Key(string source, string way, string currency) => string.Join('\u001f', source.Trim(), way.Trim(), currency.Trim()).ToUpperInvariant();
    static bool IsTotal(string value) => value.Equals("Total", StringComparison.OrdinalIgnoreCase)
        || value.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) || value.StartsWith("Total ", StringComparison.OrdinalIgnoreCase);
    static bool MatchesMonth(string name, DateTime date)
    {
        var normalized = DataValue.Normalize(name);
        return new[] { date.ToString("MMM.yyyy", CultureInfo.InvariantCulture), date.ToString("MMMM.yyyy", CultureInfo.InvariantCulture), date.ToString("yyyy.MM", CultureInfo.InvariantCulture) }
            .Any(candidate => DataValue.Normalize(candidate) == normalized);
    }
    static bool IsDate(IXLCell cell, DateTime target)
    {
        if (cell.TryGetValue<DateTime>(out var date)) return date.Date == target.Date;
        return DateTime.TryParse(cell.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Date == target.Date;
    }
    static decimal ReadAmount(IXLCell cell, string source)
    {
        var value = cell.HasFormula ? cell.CachedValue : cell.Value;
        if (value.Type == XLDataType.Number) return DataValue.Decimal(value.GetNumber().ToString("R", CultureInfo.InvariantCulture), source);
        if (cell.HasFormula) throw new InvalidDataException(source + "：公式缺少可靠数值缓存");
        return DataValue.Decimal(cell.GetString(), source);
    }
}
