using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace CashRec;

internal sealed record PreservedManualCell(decimal? Value, string SourcePath, string SourceSheet, string SourceAddress, string NumberFormat, string DisplayKey);
internal sealed record ManualStyleMapping(string SourcePath, string SourceSheet, string SourceAddress, string DestinationSheet, string DestinationAddress);

/// <summary>Imports only selected cell styles; never materializes the transaction worksheet.</summary>
internal static class ManualStyleTransfer
{
    static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static string NumberFormat(string path, string sheet, string address)
    {
        using var archive = ZipFile.OpenRead(path);
        var styles = WorkbookPackage.Read(archive, "xl/styles.xml");
        var worksheet = WorkbookPackage.Read(archive, WorkbookPackage.WorksheetPath(archive, sheet));
        var id = StyleId(worksheet, address);
        var xf = styles.Root!.Element(S + "cellXfs")!.Elements().ElementAt(id);
        var formatId = (int?)xf.Attribute("numFmtId") ?? 0;
        var custom = styles.Root.Element(S + "numFmts")?.Elements().SingleOrDefault(element => (int?)element.Attribute("numFmtId") == formatId);
        return (string?)custom?.Attribute("formatCode") ?? (formatId == 0 ? "General" : "Built-in " + formatId.ToString(CultureInfo.InvariantCulture));
    }

    internal static void Apply(string destination, IReadOnlyList<ManualStyleMapping> mappings)
    {
        if (mappings.Count == 0) return;
        using var output = ZipFile.Open(destination, ZipArchiveMode.Update);
        var targetStyles = WorkbookPackage.Read(output, "xl/styles.xml");
        var targetSheets = new Dictionary<string, XDocument>();
        foreach (var group in mappings.GroupBy(mapping => mapping.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var source = ZipFile.OpenRead(group.Key);
            var styles = WorkbookPackage.Read(source, "xl/styles.xml");
            var theme = source.GetEntry("xl/theme/theme1.xml") == null ? null : WorkbookPackage.Read(source, "xl/theme/theme1.xml");
            var sourceSheets = new Dictionary<string, XDocument>();
            var copied = new Dictionary<(string Table, int Id), int>();
            foreach (var mapping in group)
            {
                if (!sourceSheets.TryGetValue(mapping.SourceSheet, out var sheet))
                    sourceSheets[mapping.SourceSheet] = sheet = WorkbookPackage.Read(source, WorkbookPackage.WorksheetPath(source, mapping.SourceSheet));
                var id = CopyXf("cellXfs", StyleId(sheet, mapping.SourceAddress));
                var path = WorkbookPackage.WorksheetPath(output, mapping.DestinationSheet);
                if (!targetSheets.TryGetValue(path, out var target)) targetSheets[path] = target = WorkbookPackage.Read(output, path);
                var cell = target.Descendants(S + "c").SingleOrDefault(cell => (string?)cell.Attribute("r") == mapping.DestinationAddress)
                    ?? throw new InvalidDataException("无法定位人工样式目标：" + mapping.DestinationSheet + "!" + mapping.DestinationAddress);
                cell.SetAttributeValue("s", id);
            }

            int CopyXf(string table, int id)
            {
                if (copied.TryGetValue((table, id), out var saved)) return saved;
                var values = styles.Root!.Element(S + table)?.Elements().ToList() ?? throw new InvalidDataException("人工样式缺少 " + table);
                if (id < 0 || id >= values.Count) throw new InvalidDataException("人工样式编号无效");
                var value = new XElement(values[id]);
                foreach (var (attribute, component) in new[] { ("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders") })
                    value.SetAttributeValue(attribute, CopyComponent(component, (int?)value.Attribute(attribute) ?? 0));
                value.SetAttributeValue("numFmtId", CopyNumberFormat((int?)value.Attribute("numFmtId") ?? 0));
                if (table == "cellXfs" && value.Attribute("xfId") != null) value.SetAttributeValue("xfId", CopyXf("cellStyleXfs", (int)value.Attribute("xfId")!));
                var added = Add(table, value);
                copied[(table, id)] = added;
                return added;
            }
            int CopyComponent(string table, int id)
            {
                if (copied.TryGetValue((table, id), out var saved)) return saved;
                var values = styles.Root!.Element(S + table)?.Elements().ToList() ?? throw new InvalidDataException("人工样式缺少 " + table);
                if (id < 0 || id >= values.Count) throw new InvalidDataException("人工样式引用无效");
                var value = new XElement(values[id]);
                ResolveTheme(value, theme);
                var added = Add(table, value);
                copied[(table, id)] = added;
                return added;
            }
            int CopyNumberFormat(int id)
            {
                if (id < 164) return id;
                if (copied.TryGetValue(("numFmts", id), out var saved)) return saved;
                var value = styles.Root!.Element(S + "numFmts")?.Elements().SingleOrDefault(element => (int?)element.Attribute("numFmtId") == id)
                    ?? throw new InvalidDataException("人工自定义金额格式缺失");
                var table = targetStyles.Root!.Element(S + "numFmts");
                if (table == null) { table = new XElement(S + "numFmts"); targetStyles.Root.AddFirst(table); }
                var existing = table.Elements().FirstOrDefault(element => (string?)element.Attribute("formatCode") == (string?)value.Attribute("formatCode"));
                var targetId = existing == null ? table.Elements().Select(element => (int?)element.Attribute("numFmtId") ?? 163).DefaultIfEmpty(163).Max() + 1 : (int)existing.Attribute("numFmtId")!;
                if (existing == null) { var clone = new XElement(value); clone.SetAttributeValue("numFmtId", targetId); table.Add(clone); table.SetAttributeValue("count", table.Elements().Count()); }
                copied[("numFmts", id)] = targetId;
                return targetId;
            }
        }
        WorkbookPackage.Replace(output, "xl/styles.xml", targetStyles);
        foreach (var (path, sheet) in targetSheets) WorkbookPackage.Replace(output, path, sheet);

        int Add(string tableName, XElement value)
        {
            var table = targetStyles.Root!.Element(S + tableName) ?? throw new InvalidDataException("输出样式缺少 " + tableName);
            var values = table.Elements().ToList();
            var id = values.FindIndex(existing => XNode.DeepEquals(existing, value));
            if (id >= 0) return id;
            table.Add(value); table.SetAttributeValue("count", values.Count + 1); return values.Count;
        }
    }

    static int StyleId(XDocument sheet, string address)
    {
        var cell = sheet.Descendants(S + "c").SingleOrDefault(cell => (string?)cell.Attribute("r") == address);
        if (cell?.Attribute("s") is XAttribute explicitStyle) return (int)explicitStyle;
        var rowNumber = int.Parse(new string(address.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
        var row = sheet.Descendants(S + "row").SingleOrDefault(row => (int?)row.Attribute("r") == rowNumber);
        if (row?.Attribute("s") is XAttribute rowStyle) return (int)rowStyle;
        var column = 0;
        foreach (var c in address.TakeWhile(char.IsLetter)) column = column * 26 + char.ToUpperInvariant(c) - 'A' + 1;
        var inherited = sheet.Descendants(S + "col").LastOrDefault(col => (int?)col.Attribute("min") <= column && (int?)col.Attribute("max") >= column);
        return (int?)inherited?.Attribute("style") ?? 0;
    }

    static void ResolveTheme(XElement style, XDocument? theme)
    {
        if (theme == null) return;
        string[] names = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        foreach (var color in style.DescendantsAndSelf().Where(element => element.Attribute("theme") != null))
        {
            var index = (int)color.Attribute("theme")!;
            if (index < 0 || index >= names.Length) throw new InvalidDataException("人工样式主题颜色编号无效");
            var definition = theme.Descendants(A + "clrScheme").Elements(A + names[index]).Elements().SingleOrDefault();
            var rgb = (string?)definition?.Attribute("lastClr") ?? (string?)definition?.Attribute("val");
            if (rgb == null || rgb.Length != 6) throw new InvalidDataException("人工样式主题颜色不可解析");
            color.Attribute("theme")!.Remove(); color.SetAttributeValue("rgb", "FF" + rgb);
        }
        if (style.Name == S + "font" && style.Element(S + "scheme") is XElement scheme)
        {
            var kind = (string?)scheme.Attribute("val");
            if (kind is not ("major" or "minor")) return;
            var font = theme.Descendants(A + "fontScheme").Elements(A + (kind == "major" ? "majorFont" : "minorFont")).Elements(A + "latin").SingleOrDefault();
            if ((string?)font?.Attribute("typeface") is string name && name.Length > 0) { style.SetElementValue(S + "name", null); style.Element(S + "name")?.Remove(); style.AddFirst(new XElement(S + "name", new XAttribute("val", name))); scheme.Remove(); }
        }
    }
}
