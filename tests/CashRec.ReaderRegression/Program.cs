using CashRec;
using System.IO.Compression;
using System.Xml.Linq;

var directory = Directory.CreateTempSubdirectory("cashrec-reader-regression-").FullName;
var path = Path.Combine(directory, "original.xlsx");
var bank = new ReportTable("Bank");
for (var row = 2; row <= 3; row++) bank.AddRow(new()
{
    ["USD Equivalent"] = 88m,
    ["Account Number"] = "001234",
    ["DEP"] = 30m,
    ["Date"] = new DateTime(2026, 10, 25),
    ["PNL"] = new FormulaCell($"A{row}-C{row}", 58m),
    ["Branch"] = "TMA",
    ["Movement"] = null
});
WorkbookWriter.Write(path, new List<ReportTable> { bank });
Assert(BalanceWorkbookReader.HasCurrentRules(path), "Writer current-rules marker must be read");
Assert(!BalanceWorkbookReader.HasInitializedDep(path), "Pre-finalizer workbook must not claim initialized manual DEP");
var rows = BalanceWorkbookReader.ReadRows(path, "Bank");
Assert(rows.Count == 2, "Repeated records must remain for caller validation");
Assert(rows[0]["Account Number"]!.Equals("001234"), "Shared-string account must retain leading zeroes");
Assert(rows[0]["USD Equivalent"]!.Equals(88m) && rows[0]["DEP"]!.Equals(30m), "Numeric values must be decimal independent of column positions");
Assert(rows[0]["Date"]!.Equals(new DateTime(2026, 10, 25)), "Excel numeric date must retain business date");
Assert(rows[0]["PNL"]!.Equals(58m), "Engine-provided formula cache must be read");
Assert(rows[0]["Movement"] == null, "Blank financial value must remain null");
Reject(() => BalanceWorkbookReader.ReadRows(path, "Equity"), "Empty unimplemented sheet must not be a valid zero-balance snapshot");

var isolated = Copy("isolated.xlsx");
using (var archive = ZipFile.Open(isolated, ZipArchiveMode.Update))
{
    var entry = archive.GetEntry("xl/worksheets/sheet4.xml")!;
    entry.Delete();
    using (var text = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet4.xml").Open())) text.Write("<invalid-unrelated-transactions");
    Rewrite(archive, "xl/_rels/workbook.xml.rels", xml =>
    {
        var relation = xml.Descendants().Single(e => (string?)e.Attribute("Target") is "worksheets/sheet2.xml" or "/xl/worksheets/sheet2.xml");
        relation.SetAttributeValue("Target", "worksheets/moved-bank.xml");
    });
    var original = archive.GetEntry("xl/worksheets/sheet2.xml")!;
    using (var input = original.Open()) using (var output = archive.CreateEntry("xl/worksheets/moved-bank.xml").Open()) input.CopyTo(output);
    original.Delete();
    Rewrite(archive, "xl/worksheets/moved-bank.xml", xml =>
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var account = xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "B2");
        account.SetAttributeValue("t", "inlineStr");
        account.Elements().Remove();
        account.Add(new XElement(ns + "is", new XElement(ns + "t", "000777")));
        var date = xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "D2");
        date.SetAttributeValue("t", "d");
        date.Element(ns + "v")!.Value = "2026-10-25T00:00:00Z";
        var offsetDate = xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "D3");
        offsetDate.SetAttributeValue("t", "d");
        offsetDate.Element(ns + "v")!.Value = "2026-10-25T23:00:00-05:00";
    });
    Rewrite(archive, "docProps/custom.xml", xml =>
    {
        XNamespace ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        xml.Root!.Add(new XElement(ns + "property", new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
            new XAttribute("pid", "99"), new XAttribute("name", "CashRecManualDEP"), new XElement(vt + "lpwstr", "1")));
    });
}
var isolatedRows = BalanceWorkbookReader.ReadRows(isolated, "Bank");
Assert(isolatedRows[0]["Account Number"]!.Equals("000777"), "Inline-string account must retain leading zeroes");
Assert(((DateTime)isolatedRows[0]["Date"]!).Date == new DateTime(2026, 10, 25), "ISO text date must retain business day");
Assert(((DateTime)isolatedRows[1]["Date"]!).Date == new DateTime(2026, 10, 25), "Offset text date must preserve expressed business day instead of converting to local timezone");
Assert(BalanceWorkbookReader.HasInitializedDep(isolated), "Initialized manual DEP marker must be recognized");

var missing = Copy("missing-cache.xlsx");
using (var archive = ZipFile.Open(missing, ZipArchiveMode.Update)) Rewrite(archive, "xl/worksheets/sheet2.xml", xml =>
{
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "E2").Element(ns + "v")!.Remove();
});
Reject(() => BalanceWorkbookReader.ReadRows(missing, "Bank"), "Missing formula cache must not become zero");
var manualProjection = BalanceWorkbookReader.ReadRows(missing, "Bank", new[] { "Date", "Branch", "Account Number", "DEP" });
Assert(manualProjection.Count == 2 && manualProjection[0].Count == 4, "Projection must return only consumed fields and preserve records");
Assert(manualProjection[0]["DEP"]!.Equals(30m), "Numeric DEP must remain readable when unused automatic PNL cache is stale");
Assert(!manualProjection[0].ContainsKey("PNL"), "Unused formula must not be returned or validated");
Reject(() => BalanceWorkbookReader.ReadRows(missing, "Bank", new[] { "PNL" }), "Consumed PNL still requires a valid numeric cache");
Reject(() => BalanceWorkbookReader.ReadRows(missing, "Bank", new[] { "Missing Header" }), "Required projection header must be validated");

var headerOnly = Path.Combine(directory, "header-only.xlsx");
var emptyBank = new ReportTable("Bank");
emptyBank.Columns.AddRange(["Date", "Branch", "DEP"]);
WorkbookWriter.Write(headerOnly, new List<ReportTable> { emptyBank });
Assert(BalanceWorkbookReader.ReadRows(headerOnly, "Bank", new[] { "Date", "DEP" }).Count == 0, "Header-only valid schema can represent no records");
Reject(() => BalanceWorkbookReader.ReadRows(headerOnly, "Bank", new[] { "USD Equivalent" }), "Header-only sheet must still satisfy required projection schema");

var duplicate = Copy("duplicate-header.xlsx");
using (var archive = ZipFile.Open(duplicate, ZipArchiveMode.Update)) Rewrite(archive, "xl/worksheets/sheet2.xml", xml =>
{
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    var cell = xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "B1");
    cell.SetAttributeValue("t", "inlineStr");
    cell.Elements().Remove();
    cell.Add(new XElement(ns + "is", new XElement(ns + "t", "USD Equivalent")));
});
Reject(() => BalanceWorkbookReader.ReadRows(duplicate, "Bank"), "Duplicate headers must not silently overwrite");

var oldRules = Copy("old-rules.xlsx");
using (var archive = ZipFile.Open(oldRules, ZipArchiveMode.Update)) Rewrite(archive, "docProps/custom.xml", xml =>
    xml.Descendants().Single(e => (string?)e.Attribute("name") == "CashRecRules").Elements().Single().Value = "old-rules");
Assert(!BalanceWorkbookReader.HasCurrentRules(oldRules), "Old rules must not be trusted as current allocation");
Assert(!BalanceWorkbookReader.HasCompatibleManualKeys(oldRules), "Incompatible entity rules must not retain arbitrary manual keys");
var oldValuation = Copy("old-valuation.xlsx");
using (var archive = ZipFile.Open(oldValuation, ZipArchiveMode.Update)) Rewrite(archive, "docProps/custom.xml", xml =>
    xml.Descendants().Single(e => (string?)e.Attribute("name") == "CashRecValuation").Remove());
Assert(!BalanceWorkbookReader.HasCurrentRules(oldValuation), "Previous balances without current valuation marker must not be trusted");
Assert(BalanceWorkbookReader.HasCompatibleManualKeys(oldValuation), "Compatible manual key identity must survive valuation policy changes");
Console.WriteLine("PASS: marker/manual flag, reordered headers, decimal/formula caches, text identifiers, dates, blanks, duplicate records, explicit invalid-cache/header errors, targeted worksheet isolation.");
Console.WriteLine("Synthetic evidence: " + directory);

string Copy(string name) { var destination = Path.Combine(directory, name); File.Copy(path, destination); return destination; }
static void Rewrite(ZipArchive archive, string name, Action<XDocument> change)
{
    var entry = archive.GetEntry(name)!;
    XDocument xml;
    using (var stream = entry.Open()) xml = XDocument.Load(stream);
    change(xml);
    entry.Delete();
    using var output = archive.CreateEntry(name).Open();
    xml.Save(output);
}
static void Reject(Action action, string message) { try { action(); } catch (InvalidDataException) { return; } throw new Exception(message); }
static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
