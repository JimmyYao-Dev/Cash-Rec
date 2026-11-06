using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace CashRec;

internal static class WorkbookPackage
{
    internal static XDocument Read(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new InvalidDataException("工作簿缺少文件：" + path);
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XDocument.Load(reader);
    }

    internal static string WorksheetPath(ZipArchive archive, string name)
    {
        var sheet = Read(archive, "xl/workbook.xml").Descendants().SingleOrDefault(element => element.Name.LocalName == "sheet" && (string?)element.Attribute("name") == name)
            ?? throw new InvalidDataException("缺少工作表：" + name);
        var id = sheet.Attributes().SingleOrDefault(attribute => attribute.Name.LocalName == "id")?.Value;
        var relation = Read(archive, "xl/_rels/workbook.xml.rels").Descendants().SingleOrDefault(element => element.Name.LocalName == "Relationship" && (string?)element.Attribute("Id") == id)
            ?? throw new InvalidDataException(name + " 工作表关系缺失");
        if ((string?)relation.Attribute("TargetMode") == "External" || !((string?)relation.Attribute("Type") ?? "").EndsWith("/worksheet", StringComparison.Ordinal))
            throw new InvalidDataException(name + " 不读取外部或非工作表关系");
        var uri = new Uri(new Uri("https://cashrec.invalid/xl/workbook.xml"), (string?)relation.Attribute("Target") ?? "");
        if (uri.Host != "cashrec.invalid") throw new InvalidDataException(name + " 工作表关系指向外部位置");
        return Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
    }

    internal static void Replace(ZipArchive archive, string path, XDocument xml)
    {
        archive.GetEntry(path)?.Delete();
        using var stream = archive.CreateEntry(path).Open();
        xml.Save(stream);
    }
}
