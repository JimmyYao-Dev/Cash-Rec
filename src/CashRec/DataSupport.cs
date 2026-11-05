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
            var key = row.Keys.FirstOrDefault(k => Normalize(k) == Normalize(name));
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
    public static IEnumerable<Dictionary<string, string>> Read(string path, string? headerMarker = null)
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
                if (headers.Distinct(StringComparer.OrdinalIgnoreCase).Count() != headers.Length)
                    throw new InvalidDataException($"{Path.GetFileName(path)}：表头重复");
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
        if (!values.TryGetValue(currency.Trim(), out var rate) || rate <= 0)
            throw new InvalidDataException($"汇率：{currency} 的 Units per USD 缺失或不大于零");
        return rate;
    }
    public decimal Convert(decimal amount, string currency) => amount / Get(currency);
    public static DailyRates Load(string path)
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
            var rates = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var dataRow in sheet.RowsUsed().Where(r => r.RowNumber() > row.RowNumber()))
            {
                var currency = dataRow.Cell(currencyCell.Address.ColumnNumber).GetString().Trim();
                if (currency.Length == 0) continue;
                var text = dataRow.Cell(rateCell.Address.ColumnNumber).GetFormattedString();
                var rate = DataValue.Decimal(text, "汇率 " + currency);
                if (rate <= 0) throw new InvalidDataException($"汇率：{currency} 必须大于零");
                if (rates.TryGetValue(currency, out var previous) && previous != rate) throw new InvalidDataException($"汇率：{currency} 有冲突记录");
                rates[currency] = rate;
            }
            return new DailyRates(rates);
        }
        throw new InvalidDataException("汇率：未找到 Currency 与 Units per USD 表头");
    }
}
