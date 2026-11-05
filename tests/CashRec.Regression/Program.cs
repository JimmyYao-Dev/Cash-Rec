using CashRec;
using ClosedXML.Excel;
using System.IO.Compression;
using System.Xml.Linq;

var directory = Directory.CreateTempSubdirectory("cashrec-regression-").FullName;
var ratePath = Path.Combine(directory, "rates.xlsx");
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.AddWorksheet("Rates");
    sheet.Cell(1, 1).Value = "Symb";
    sheet.Cell(1, 2).Value = "Currency";
    sheet.Cell(1, 3).Value = "Units per USD";
    sheet.Cell(2, 1).Value = "USD";
    sheet.Cell(2, 2).Value = "US Dollar";
    sheet.Cell(2, 3).Value = 1.25;
    sheet.Cell(2, 3).Style.NumberFormat.Format = "0.0";
    workbook.SaveAs(ratePath);
}
var rates = DailyRates.Load(ratePath);
Assert(rates.Get("USD") == 1.25m, "FX must read raw 1.25, not displayed 1.3 or currency description");
Assert(rates.Convert(100m, "USD") == 80m, "FX divisor must retain source precision");

var selectedDateFolder = Path.Combine(directory, "2026.10.25");
Directory.CreateDirectory(selectedDateFolder);
var source = Path.Combine(directory, "NAB AU.csv");
File.WriteAllText(source, "Account Number,Currency,EOD balance\n0001,USD,100\n");
try
{
    WorkspaceInputs.Import(selectedDateFolder, new DateTime(2026, 10, 26), source);
    throw new Exception("Wrong-date import unexpectedly succeeded");
}
catch (InvalidDataException) { }
Assert(!Directory.EnumerateFileSystemEntries(selectedDateFolder).Any(), "Wrong-date import must not create files, backups or directories");
Assert(WorkspaceInputs.Inspect(selectedDateFolder, new DateTime(2026, 10, 26)).Diagnostics.Any(d => d.Source == "日期" && d.Severity == "Error"), "Preflight must retain explicit date mismatch diagnostic");

var table = new ReportTable("Equity");
table.AddRow(new() { ["Movement"] = 8m, ["DEP"] = 30m, ["PNL"] = new FormulaCell("A2-B2", -22m) });
var outputPath = Path.Combine(directory, "cache.xlsx");
WorkbookWriter.Write(outputPath, new List<ReportTable> { table });
using (var archive = ZipFile.OpenRead(outputPath))
{
    using var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
    var xml = XDocument.Load(stream);
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    var cell = xml.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == "C2");
    Assert((string?)cell.Element(ns + "f") == "A2-B2", "PNL formula must remain editable");
    Assert((string?)cell.Element(ns + "v") == "-22", "PNL cache must contain engine-calculated value");
}
Console.WriteLine("PASS: FX raw precision; wrong-date import protection; formula numeric cache.");
Console.WriteLine("Synthetic evidence: " + directory);

static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
