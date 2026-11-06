using System.Globalization;
using System.Text;
using Microsoft.VisualBasic.FileIO;
using ClosedXML.Excel;

namespace CashRec;

public static class DataValue
{
    public static string Normalize(string value) => new(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    public static string Get(IReadOnlyDictionary<string, string> row, params string[] names)
    {
        foreach (var name in names)
        {
            if (row.TryGetValue(name, out var exact)) return exact.Trim();
            var normalizedName = Normalize(name);
            var key = row.Keys.FirstOrDefault(k => Normalize(k) == normalizedName);
            if (key != null) return row[key].Trim();
        }
        return "";
    }
    public static decimal Decimal(string value, string source)
    {
        var text = value.Trim();
        if (text.StartsWith('(') && text.EndsWith(')')) text = "-" + text[1..^1];
        if (decimal.TryParse(text, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out var number)) return number;
        throw new InvalidDataException($"{source}：无法识别金额 {value}");
    }
    public static string Account(string value) => new(value.Where(c => !char.IsWhiteSpace(c) && c != '-').ToArray());
}

public static class CsvData
{
    public static IEnumerable<Dictionary<string, string>> Read(string path, string? headerMarker = null, IReadOnlyCollection<string>? requiredHeaders = null)
    {
        using var parser = new TextFieldParser(path, Encoding.UTF8, true);
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters(",");
        parser.HasFieldsEnclosedInQuotes = true;
        string[]? headers = null;
        while (!parser.EndOfData)
        {
            var values = parser.ReadFields();
            if (values == null || values.All(string.IsNullOrWhiteSpace)) continue;
            if (headers == null)
            {
                if (headerMarker != null && !values.Any(v => DataValue.Normalize(v) == DataValue.Normalize(headerMarker))) continue;
                headers = values.Select(x => x.Trim().Trim('\uFEFF')).ToArray();
                var normalizedHeaders = headers.Select(DataValue.Normalize).ToHashSet(StringComparer.Ordinal);
                if (normalizedHeaders.Count != headers.Length)
                    throw new InvalidDataException($"{Path.GetFileName(path)}：表头重复");
                var missing = requiredHeaders?.Where(name => !normalizedHeaders.Contains(DataValue.Normalize(name))).ToArray() ?? [];
                if (missing.Length > 0)
                    throw new InvalidDataException($"{Path.GetFileName(path)}：缺少表头 {string.Join(", ", missing)}");
                continue;
            }
            if (values.Length > headers.Length) throw new InvalidDataException($"{Path.GetFileName(path)}：行字段多于表头");
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Length; i++) row[headers[i]] = i < values.Length ? values[i] : "";
            yield return row;
        }
        if (headers == null) throw new InvalidDataException($"{Path.GetFileName(path)}：未找到表头");
    }
}

public sealed class DailyRates
{
    readonly Dictionary<string, decimal> values;
    public DailyRates(Dictionary<string, decimal> values) => this.values = new(values, StringComparer.OrdinalIgnoreCase);
    public decimal Get(string currency)
    {
        if (currency.Trim().Equals("USDT", StringComparison.OrdinalIgnoreCase)) return 1m;
        if (!values.TryGetValue(currency.Trim(), out var rate) || rate <= 0)
            throw new InvalidDataException($"汇率：{currency} 的 Units per USD 缺失或不大于零");
        return rate;
    }
    public decimal Convert(decimal amount, string currency) => amount / Get(currency);
    public static DailyRates Load(string path, DateTime? businessDate = null)
    {
        using var workbook = new XLWorkbook(path);
        foreach (var sheet in workbook.Worksheets)
        foreach (var row in sheet.RowsUsed().Take(30))
        {
            var cells = row.CellsUsed().ToList();
            var currencyCell = cells.FirstOrDefault(c => new[] { "symb", "symbol" }.Contains(DataValue.Normalize(c.GetString())))
                ?? cells.FirstOrDefault(c => new[] { "currency", "currencycode", "curr", "code" }.Contains(DataValue.Normalize(c.GetString())));
            var rateCell = cells.FirstOrDefault(c => DataValue.Normalize(c.GetString()) == "unitsperusd");
            if (currencyCell == null || rateCell == null) continue;
            var dateCell = cells.FirstOrDefault(c => DataValue.Normalize(c.GetString()) == "date");
            var dataRows = sheet.RowsUsed().Where(r => r.RowNumber() > row.RowNumber()).ToList();
            if (dateCell != null)
            {
                var datedRows = dataRows.Select(data => (Row: data, Date: ReadDate(data.Cell(dateCell.Address.ColumnNumber), path))).ToList();
                var dates = datedRows.Where(data => data.Date != null).Select(data => data.Date!.Value).Distinct().ToList();
                if (businessDate == null && dates.Count > 1)
                    throw new InvalidDataException("汇率文件包含多个日期，必须明确业务日期，不能混用历史汇率");
                var selectedDate = businessDate?.Date ?? dates.SingleOrDefault();
                dataRows = datedRows.Where(data => data.Date == selectedDate).Select(data => data.Row).ToList();
                if (dataRows.Count == 0)
                    throw new InvalidDataException($"汇率文件没有业务日期 {selectedDate:yyyy.MM.dd} 的数据，不能使用其他日期");
            }
            var rates = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var dataRow in dataRows)
            {
                var symbol = dataRow.Cell(currencyCell.Address.ColumnNumber);
                var symbolValue = symbol.HasFormula ? symbol.CachedValue : symbol.Value;
                if (symbol.HasFormula && (symbolValue.Type != XLDataType.Text || string.IsNullOrWhiteSpace(symbolValue.GetText())))
                    throw new InvalidDataException($"汇率 {sheet.Name}!{symbol.Address}：Symb公式缺少有效文本缓存");
                var currency = (symbol.HasFormula ? symbolValue.GetText() : symbol.GetString()).Trim().ToUpperInvariant();
                if (currency.Length == 0) continue;
                if (currency == "USDT") { rates[currency] = 1m; continue; }
                var rate = ExcelNumbers.Read(dataRow.Cell(rateCell.Address.ColumnNumber), "汇率 " + currency);
                if (rate <= 0) throw new InvalidDataException($"汇率：{currency} 必须大于零");
                if (rates.TryGetValue(currency, out var previous) && previous != rate) throw new InvalidDataException($"汇率：{currency} 有冲突记录");
                rates[currency] = rate;
            }
            if (rates.Count == 0) throw new InvalidDataException("汇率文件没有可用的币种记录");
            rates.TryAdd("USD", 1m);
            rates.TryAdd("USC", 100m);
            rates["USDT"] = 1m;
            return new DailyRates(rates);
        }
        throw new InvalidDataException("汇率：未找到 Currency 与 Units per USD 表头");
    }

    static DateTime? ReadDate(IXLCell cell, string path)
    {
        var value = cell.HasFormula ? cell.CachedValue : cell.Value;
        if (cell.HasFormula && value.Type is XLDataType.Blank or XLDataType.Error)
            throw new InvalidDataException($"{Path.GetFileName(path)} {cell.Address}：日期公式缺少有效缓存");
        if (value.Type == XLDataType.Blank) return null;
        if (value.Type == XLDataType.DateTime) return value.GetDateTime().Date;
        if (value.Type == XLDataType.Number)
        {
            try { return DateTime.FromOADate(value.GetNumber()).Date; }
            catch (ArgumentException) { throw new InvalidDataException($"{Path.GetFileName(path)} {cell.Address}：无效汇率日期"); }
        }
        if (value.Type == XLDataType.Text && DateTimeOffset.TryParseExact(value.GetText().Trim(),
            ["yyyy.MM.dd", "yyyy-MM-dd", "yyyy/MM/dd", "yyyy-MM-dd HH:mm:ss", "s", "O"],
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return date.Date;
        if (value.Type == XLDataType.Text && string.IsNullOrWhiteSpace(value.GetText())) return null;
        throw new InvalidDataException($"{Path.GetFileName(path)} {cell.Address}：无法识别汇率日期 {value}");
    }
}
