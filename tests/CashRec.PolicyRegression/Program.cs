using CashRec;
using ClosedXML.Excel;
using System.IO.Compression;
using System.Xml.Linq;

try
{
var directory = Directory.CreateTempSubdirectory("cashrec-policy-regression-").FullName;
var current = new DateTime(2026, 10, 25);
var prior = new DateTime(2026, 10, 22);
var ratesPath = Path.Combine(directory, "historical-rates.xlsx");
using (var book = new XLWorkbook())
{
    var sheet = book.AddWorksheet("Rates");
    string[] headers = ["Date", "Symb", "Currency", "Units per USD"];
    for (var column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).Value = headers[column];
    object[][] rows = [
        ["2026.10.22", "", "AUD Australian Dollar", 1.5],
        ["2026.10.25", "", "AUD Australian Dollar", 1.25],
        ["2026.10.26", "", "AUD Australian Dollar", "invalid irrelevant rate"],
        ["2026.10.25", "USDT", "USDT Tether", 0.97]
    ];
    for (var row = 0; row < rows.Length; row++)
    {
        sheet.Cell(row + 2, 1).Value = (string)rows[row][0];
        sheet.Cell(row + 2, 2).Value = (string)rows[row][1];
        sheet.Cell(row + 2, 3).Value = (string)rows[row][2];
        if (rows[row][3] is double number) sheet.Cell(row + 2, 4).Value = number;
        else sheet.Cell(row + 2, 4).Value = (string)rows[row][3];
        if (row < 3) sheet.Cell(row + 2, 2).FormulaA1 = $"LEFT(C{row + 2},3)";
    }
    sheet.Cell(3, 4).Style.NumberFormat.Format = "0.0";
    book.SaveAs(ratesPath, new ClosedXML.Excel.SaveOptions { EvaluateFormulasBeforeSaving = true });
}
using (var archive = ZipFile.Open(ratesPath, ZipArchiveMode.Update))
{
    var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
    XDocument xml;
    using (var stream = entry.Open()) xml = XDocument.Load(stream);
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    xml.Descendants(ns + "c").Single(e => (string?)e.Attribute("r") == "B4").Element(ns + "v")!.Remove();
    entry.Delete();
    using var output = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
    xml.Save(output);
}
var daily = DailyRates.Load(ratesPath, current);
Assert(daily.Get("AUD") == 1.25m && daily.Convert(100m, "AUD") == 80m, "Selected-day cached Symb/raw rate must retain precision");
Assert(DailyRates.Load(ratesPath, prior).Get("AUD") == 1.5m, "Prior date must use its own rate, not current or latest rate");
Assert(daily.Get("USD") == 1m && daily.Get("USC") == 100m, "Missing base/cents unit entries use canonical identities");
Assert(daily.Get("USDT") == 1m, "Explicit non-1 USDT market input must not replace the fixed business rate");
Assert(new DailyRates(new() { ["USDT"] = 0.91m }).Convert(100m, "USDT") == 100m, "USDT business identity applies to every rate context");
Assert(DailyRates.Load(ratesPath, prior).Get("USDT") == 1m, "Missing USDT source rate still has fixed business valuation");
Reject(() => DailyRates.Load(ratesPath), "多个日期", "History requires explicit business date");
Reject(() => DailyRates.Load(ratesPath, current.AddDays(-1)), "2026.10.24", "Missing exact rate date must not borrow another date");
Reject(() => DailyRates.Load(ratesPath, current.AddDays(1)), "缓存", "Consumed uncached Symb formula remains an error");

var mappingPath = Path.Combine(directory, "account_mapping.xlsx");
using (var book = new XLWorkbook())
{
    var bank = book.AddWorksheet("BANK MAPPING");
    string[] bankHeaders = ["Account Number", "Source File", "Company", "Account Type", "Client Type", "Settle Type"];
    for (var c = 0; c < bankHeaders.Length; c++) bank.Cell(1, c + 1).Value = bankHeaders[c];
    var psp = book.AddWorksheet("PSP MAPPING");
    string[] pspHeaders = ["Source", "PSP", "Currency", "Company", "Branch"];
    string[] pspValues = ["PSPAU", "KnownGateway", "USDT", "TMC", "TMA"];
    for (var c = 0; c < pspHeaders.Length; c++) { psp.Cell(1, c + 1).Value = pspHeaders[c]; psp.Cell(2, c + 1).Value = pspValues[c]; }
    book.SaveAs(mappingPath);
}
var pspPath = Path.Combine(directory, "psp.xlsx");
WritePsp(pspPath);
var context = new RunContext(current, directory, new() { ["PSP"] = pspPath }, mappingPath, daily);
var zero = new PspAdapter().Process(context);
Assert(zero.Rows.Count == 1 && zero.Rows[0]["Branch"]!.Equals("TMA"), "Blank PSP row must remain in known source/entity");
Assert(zero.Rows[0]["Amount"]!.Equals(0m) && zero.Rows[0]["USD Equivalent"]!.Equals(0m), "True blank PSP Amount becomes numeric zero");
using (var book = new XLWorkbook(pspPath))
{
    Reject(() => ExcelNumbers.Read(book.Worksheet(1).Cell(3, 5), "Other input"), "金额", "Other blank financial inputs remain strict");
    book.Worksheet(1).Cell(3, 5).FormulaA1 = "1+1";
    book.Save();
}
Reject(() => new PspAdapter().Process(context), "缓存", "Uncached PSP formula must not become zero");
using (var book = new XLWorkbook(pspPath)) { book.Worksheet(1).Cell(3, 5).Value = "invalid amount"; book.Save(); }
Reject(() => new PspAdapter().Process(context), "金额", "Invalid PSP text must not become zero");

var workspace = Path.Combine(directory, "workspace");
Directory.CreateDirectory(Path.Combine(workspace, "config"));
File.Copy(mappingPath, Path.Combine(workspace, "config", "account_mapping.xlsx"));
foreach (var day in new[] { prior, current })
{
    var needed = Path.Combine(workspace, day.ToString("yyyy.MM.dd"), "Needed files");
    foreach (var definition in WorkspaceInputs.Definitions) Directory.CreateDirectory(Path.GetDirectoryName(Path.Combine(needed, definition.RelativePath))!);
    File.WriteAllText(Path.Combine(needed, "MT4/Risk/Accounting Equity Report.csv"),
        "dt_report,report_jurisdiction,is_ib,positive_equity,currency\n2026-10-25,ASIC,true," + (day == prior ? "150" : "125") + ",AUD\n");
    File.WriteAllText(Path.Combine(needed, "MT4/MT4/Transaction Summary Report.csv"), "date,jurisdiction,accounting_comment,comment,fund_amount,trading_server_currency\n");
    foreach (var name in new[] { "NAB AU.csv", "NAB GL.csv" }) File.WriteAllText(Path.Combine(needed, "Bank", name), "Account Name,Account Number,Currency,EOD balance\n");
    WritePsp(Path.Combine(needed, "PSP+BANK/PSP Balance Summary.xlsx"));
    using (var book = new XLWorkbook())
    {
        var sheet = book.AddWorksheet("BANK BD");
        string[] headers = ["Comment", "Rational", "Currency", "Company", "Amount"];
        for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
        book.SaveAs(Path.Combine(needed, "PSP+BANK/BANK BD.xlsx"));
    }
    File.Copy(ratesPath, Path.Combine(needed, "Rate/VFSC Rate.xlsx"));
}
var result = ReportRunner.Run(workspace, current);
Assert(result.IsComplete, string.Join(";", result.Diagnostics.Select(d => d.Message)));
var equity = BalanceWorkbookReader.ReadRows(result.OutputPath!, "Equity").Single();
Assert(equity["USD Equivalent"]!.Equals(100m) && equity["Prev USD Equivalent"]!.Equals(100m) && equity["Movement"]!.Equals(0m), "Production current and prior contexts must pass their own dates to history loader");
Assert(BalanceWorkbookReader.HasCurrentRules(result.OutputPath!), "New output must declare current valuation policy");
Console.WriteLine("PASS: dated FX and cached symbols; fixed USDT; base unit defaults; true blank PSP zero; strict consumed caches/text; current/prior production date propagation.");
Console.WriteLine("Synthetic evidence: " + directory);

void WritePsp(string file)
{
    using var book = new XLWorkbook();
    var sheet = book.AddWorksheet("OCT.2026");
    sheet.Cell(1, 4).Value = prior;
    sheet.Cell(1, 5).Value = current;
    string[] headers = ["Source", "WAY", "CURR", "Amount", "Amount"];
    for (var c = 0; c < headers.Length; c++) sheet.Cell(2, c + 1).Value = headers[c];
    sheet.Cell(3, 1).Value = "PSPAU";
    sheet.Cell(3, 2).Value = "KnownGateway";
    sheet.Cell(3, 3).Value = "USDT";
    // Both dates' Amount cells are true blanks; no formula or invalid text is substituted.
    book.SaveAs(file);
}
static void Reject(Action action, string text, string message) { try { action(); } catch (InvalidDataException error) { if (error.Message.Contains(text)) return; throw new Exception(message + ": " + error.Message); } throw new Exception(message); }
static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
}
catch (Exception error)
{
    Console.Error.WriteLine(error);
    Environment.Exit(1);
}
