using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace CashRec;

/// <summary>Reads selected balance sheets without loading transaction worksheets.</summary>
public static class BalanceWorkbookReader
{
    static readonly HashSet<string> FinancialHeaders = new(StringComparer.Ordinal)
    {
        "amount", "eodbalance", "usdequivalent", "prevusdequivalent", "previous", "movement", "dep", "pnl", "unitsperusd"
    };

    public static bool HasCurrentRules(string path) => HasCompatibleManualKeys(path) && HasProperty(path, "CashRecValuation", ReportPolicy.Valuation);
    public static bool HasCompatibleManualKeys(string path) => HasProperty(path, "CashRecRules", ReportPolicy.EntityRules);
    public static bool HasInitializedDep(string path) => HasProperty(path, "CashRecManualDEP", "1");

    static bool HasProperty(string path, string name, string expected) => ReadPackage(path, archive =>
    {
        var entry = archive.GetEntry("docProps/custom.xml");
        if (entry == null) return false;
        using var stream = entry.Open();
        var xml = LoadXml(stream);
        var properties = xml.Descendants().Where(e => e.Name.LocalName == "property" && (string?)e.Attribute("name") == name).ToList();
        return properties.Count == 1 && properties[0].Elements().Count() == 1 && properties[0].Elements().Single().Value == expected;
    });

    public static List<Dictionary<string, object?>> ReadRows(string path, string sheetName) =>
        ReadPackage(path, archive => ReadRowsCore(archive, path, sheetName, null));

    public static List<Dictionary<string, object?>> ReadRows(string path, string sheetName, IReadOnlyCollection<string> selectedColumns) =>
        ReadPackage(path, archive => ReadRowsCore(archive, path, sheetName, selectedColumns));

    static List<Dictionary<string, object?>> ReadRowsCore(ZipArchive archive, string path, string sheetName, IReadOnlyCollection<string>? selectedColumns)
    {
        var workbook = ReadXml(archive, "xl/workbook.xml");
        var sheet = workbook.Descendants().SingleOrDefault(e => e.Name.LocalName == "sheet" && (string?)e.Attribute("name") == sheetName)
            ?? throw new InvalidDataException($"{Path.GetFileName(path)}：缺少 {sheetName} 工作表");
        var relationshipId = sheet.Attributes().SingleOrDefault(a => a.Name.LocalName == "id")?.Value
            ?? throw new InvalidDataException($"{sheetName}：工作表关系缺失");
        var relationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var relationship = relationships.Descendants().SingleOrDefault(e => e.Name.LocalName == "Relationship" && (string?)e.Attribute("Id") == relationshipId)
            ?? throw new InvalidDataException($"{sheetName}：无法定位工作表文件");
        if (((string?)relationship.Attribute("TargetMode"))?.Equals("External", StringComparison.OrdinalIgnoreCase) == true)
            throw new InvalidDataException($"{sheetName}：不读取外部工作表关系");
        if (!((string?)relationship.Attribute("Type") ?? "").EndsWith("/worksheet", StringComparison.Ordinal))
            throw new InvalidDataException($"{sheetName}：关系不是工作表");
        var target = (string?)relationship.Attribute("Target") ?? throw new InvalidDataException($"{sheetName}：工作表路径缺失");
        var targetUri = new Uri(new Uri("https://cashrec.invalid/xl/workbook.xml"), target);
        if (targetUri.Host != "cashrec.invalid") throw new InvalidDataException($"{sheetName}：工作表关系指向外部位置");
        var worksheetPath = Uri.UnescapeDataString(targetUri.AbsolutePath).TrimStart('/');
        var worksheet = ReadXml(archive, worksheetPath);
        var date1904 = workbook.Descendants().Where(e => e.Name.LocalName == "workbookPr")
            .Any(e => (string?)e.Attribute("date1904") is "1" or "true");
        var sheetRows = worksheet.Descendants().Where(e => e.Name.LocalName == "sheetData").Elements()
            .Where(e => e.Name.LocalName == "row").ToList();
        var rows = new List<Dictionary<string, object?>>();
        Dictionary<int, string>? headers = null;
        var headerIndex = -1;
        for (var index = 0; index < sheetRows.Count; index++)
        {
            var cells = IndexedCells(sheetRows[index]);
            var headerStrings = ReadSharedStrings(archive, cells.Values);
            var names = cells.Select(pair => (pair.Key, Value: ReadCell(pair.Value, headerStrings, "", sheetName))).Where(pair => pair.Value != null).ToList();
            if (names.Count == 0) continue;
            headers = new Dictionary<int, string>();
            var normalized = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (column, value) in names)
            {
                if (value is not string text || string.IsNullOrWhiteSpace(text))
                    throw new InvalidDataException($"{sheetName}：表头必须为非空文本，列 {column}");
                text = text.Trim();
                if (!normalized.Add(DataValue.Normalize(text))) throw new InvalidDataException($"{sheetName}：重复表头 {text}");
                headers.Add(column, text);
            }
            headerIndex = index;
            break;
        }
        if (headers == null) throw new InvalidDataException($"{sheetName}：没有可识别的表头行");
        if (selectedColumns != null) headers = ProjectHeaders(headers, selectedColumns, sheetName);
        var dataRows = sheetRows.Skip(headerIndex + 1).Select(IndexedCells).ToList();
        var dataCells = dataRows.SelectMany(row => row.Where(pair => selectedColumns == null || headers.ContainsKey(pair.Key)).Select(pair => pair.Value));
        var sharedStrings = ReadSharedStrings(archive, dataCells);
        for (var index = 0; index < dataRows.Count; index++)
        {
            var recordNumber = headerIndex + index + 2;
            var result = headers.Values.ToDictionary(header => header, _ => (object?)null, StringComparer.OrdinalIgnoreCase);
            var hasValue = false;
            foreach (var (column, cell) in dataRows[index])
            {
                if (selectedColumns != null && !headers.ContainsKey(column)) continue;
                var header = headers.GetValueOrDefault(column) ?? "";
                var value = ReadCell(cell, sharedStrings, header, sheetName);
                if (value == null) continue;
                if (header.Length == 0) throw new InvalidDataException($"{sheetName}：第 {recordNumber} 行的列 {column} 有值但没有表头");
                if (DataValue.Normalize(header) == "date") value = ReadDate(value, date1904, sheetName, recordNumber);
                result[header] = value;
                hasValue = true;
            }
            if (hasValue) rows.Add(result);
        }
        return rows;
    }

    static Dictionary<int, string> ProjectHeaders(Dictionary<int, string> headers, IReadOnlyCollection<string> selectedColumns, string sheetName)
    {
        if (selectedColumns.Count == 0) throw new InvalidDataException($"{sheetName}：未指定需要读取的字段");
        var actualColumns = headers.ToDictionary(pair => DataValue.Normalize(pair.Value), pair => pair.Key, StringComparer.Ordinal);
        var projected = new Dictionary<int, string>();
        foreach (var name in selectedColumns)
        {
            if (!actualColumns.TryGetValue(DataValue.Normalize(name), out var column))
                throw new InvalidDataException($"{sheetName}：缺少所需表头 {name}");
            if (!projected.TryAdd(column, name)) throw new InvalidDataException($"{sheetName}：重复请求字段 {name}");
        }
        return projected;
    }

    static FileStream OpenFile(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    static T ReadPackage<T>(string path, Func<ZipArchive, T> read)
    {
        try
        {
            using var file = OpenFile(path);
            using var archive = new ZipArchive(file, ZipArchiveMode.Read);
            return read(archive);
        }
        catch (Exception error) when (error is XmlException or InvalidOperationException or ArgumentException or OverflowException)
        { throw new InvalidDataException($"{Path.GetFileName(path)}：工作簿结构无效：{error.Message}", error); }
    }
    static XDocument LoadXml(Stream stream)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XDocument.Load(reader);
    }
    static XDocument ReadXml(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName) ?? throw new InvalidDataException("工作簿缺少文件：" + entryName);
        using var stream = entry.Open();
        return LoadXml(stream);
    }

    static Dictionary<int, XElement> IndexedCells(XElement row)
    {
        var result = new Dictionary<int, XElement>();
        var previous = 0;
        foreach (var cell in row.Elements().Where(e => e.Name.LocalName == "c"))
        {
            var address = (string?)cell.Attribute("r");
            var column = address == null ? previous + 1 : Column(address);
            if (!result.TryAdd(column, cell)) throw new InvalidDataException("工作表同一行有重复单元格：" + address);
            previous = column;
        }
        return result;
    }
    static int Column(string address)
    {
        var result = 0;
        foreach (var character in address.TakeWhile(char.IsLetter))
            result = checked(result * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        if (result == 0) throw new InvalidDataException("无效的单元格地址：" + address);
        return result;
    }

    static Dictionary<int, string> ReadSharedStrings(ZipArchive archive, IEnumerable<XElement> cells)
    {
        var needed = new HashSet<int>();
        foreach (var cell in cells.Where(e => (string?)e.Attribute("t") == "s"))
        {
            var value = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v")?.Value;
            if (value == null) continue;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
                throw new InvalidDataException("无效的共享文本索引：" + value);
            needed.Add(index);
        }
        var values = new Dictionary<int, string>();
        if (needed.Count == 0) return values;
        var entry = archive.GetEntry("xl/sharedStrings.xml") ?? throw new InvalidDataException("工作簿缺少共享文本");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var current = 0;
        var maximum = needed.Max();
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si") continue;
            using (var item = reader.ReadSubtree())
            {
                if (needed.Contains(current)) values.Add(current, Text(XElement.Load(item)));
                else while (item.Read()) { }
            }
            if (current++ >= maximum) break;
        }
        if (values.Count != needed.Count) throw new InvalidDataException("工作簿共享文本索引超出范围");
        return values;
    }

    static string Text(XElement element) => string.Concat(element.Descendants().Where(e => e.Name.LocalName == "t"
        && !e.Ancestors().Any(a => a.Name.LocalName == "rPh")).Select(e => e.Value));

    static object? ReadCell(XElement cell, IReadOnlyDictionary<int, string> strings, string header, string sheetName)
    {
        var type = (string?)cell.Attribute("t") ?? "n";
        var raw = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v")?.Value;
        var formula = cell.Elements().Any(e => e.Name.LocalName == "f");
        var financial = FinancialHeaders.Contains(DataValue.Normalize(header));
        if (formula && financial && (type != "n" || string.IsNullOrWhiteSpace(raw)
            || !decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _)))
            throw new InvalidDataException($"{sheetName} {cell.Attribute("r")?.Value}：{header} 公式缺少有效数值缓存");
        if (type == "inlineStr") return cell.Elements().FirstOrDefault(e => e.Name.LocalName == "is") is { } inline && Text(inline) is { Length: > 0 } inlineText ? inlineText : null;
        if (raw == null || raw.Length == 0) return null;
        if (type == "s")
        {
            if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var index) || !strings.TryGetValue(index, out var text))
                throw new InvalidDataException("无效的共享文本索引：" + raw);
            return text.Length == 0 ? null : text;
        }
        if (type is "str" or "d") return raw;
        if (type == "b") return raw switch { "1" => true, "0" => false, _ => throw new InvalidDataException("无效的布尔单元格：" + raw) };
        if (type == "e") throw new InvalidDataException($"{sheetName} {cell.Attribute("r")?.Value}：单元格错误 {raw}");
        if (type != "n" || !decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            throw new InvalidDataException($"{sheetName} {cell.Attribute("r")?.Value}：无效数值 {raw}");
        return number;
    }

    static DateTime ReadDate(object value, bool date1904, string sheet, int row)
    {
        if (value is decimal serial)
        {
            try { return DateTime.FromOADate((double)serial + (date1904 ? 1462 : 0)); }
            catch (ArgumentException) { throw new InvalidDataException($"{sheet} 第 {row} 行：无效日期数值 {serial}"); }
        }
        if (value is string text && DateTimeOffset.TryParseExact(text.Trim(),
            ["yyyy.MM.dd", "yyyy-MM-dd", "yyyy/MM/dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ssK", "yyyy-MM-ddTHH:mm:ss.FFFFFFFK", "O"],
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return date.DateTime;
        throw new InvalidDataException($"{sheet} 第 {row} 行：无法识别完整业务日期 {value}");
    }
}
