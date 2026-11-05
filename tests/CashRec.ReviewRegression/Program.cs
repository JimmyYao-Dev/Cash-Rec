using CashRec;
using ClosedXML.Excel;

var directory = Directory.CreateTempSubdirectory("cashrec-review-regression-").FullName;
var equity = Path.Combine(directory, "Accounting Equity Report.csv");
var mt4 = Path.Combine(directory, "Transaction Summary Report.csv");
var au = Path.Combine(directory, "NAB AU.csv");
var gl = Path.Combine(directory, "NAB GL.csv");
var mapping = Path.Combine(directory, "mapping.xlsx");
var context = new RunContext(new DateTime(2026, 10, 25), directory,
    new() { ["Equity"] = equity, ["Transactions"] = mt4, ["BankAU"] = au, ["BankGL"] = gl },
    mapping, new DailyRates(new() { ["USD"] = 1m }));
WriteMapping("TMC", "Clients", "Retail", "NA");
File.WriteAllText(gl, "Account Name,Account Number,Currency,EOD balance\n");
File.WriteAllText(equity, "dt_report,report_jurisdiction\n");
Reject(() => new EquityAdapter().Process(context), "positive_equity", "Malformed header-only Equity must fail");
File.WriteAllText(mt4, "date,jurisdiction\n");
Reject(() => new Mt4TransactionAdapter().Process(context), "fund_amount", "Malformed header-only MT4 must fail");
File.WriteAllText(au, "Account Number,Currency\n");
Reject(() => new BankAdapter().Process(context), "EOD balance", "Malformed header-only NAB must fail");
var duplicate = Path.Combine(directory, "duplicate.csv");
File.WriteAllText(duplicate, "fund_amount,fund amount\n");
Reject(() => CsvData.Read(duplicate).ToList(), "重复", "Normalized duplicate header must fail before data exists");

File.WriteAllText(equity, "dt_report,report_jurisdiction,is_ib,positive_equity,currency\n");
var emptyEquity = new EquityAdapter().Process(context);
Assert(emptyEquity.Rows.Count == 0 && emptyEquity.Diagnostics.Any(d => d.Severity == "Warning" && d.Message.Contains("没有权益")), "Valid empty Equity requires explicit zero-record diagnostic, not a schema error");
File.WriteAllText(mt4, "date,jurisdiction,accounting_comment,comment,fund_amount,trading_server_currency\n");
Assert(new Mt4TransactionAdapter().Process(context).Diagnostics.Any(d => d.Message.Contains("没有MT4")), "Valid empty MT4 requires explicit zero-record diagnostic");
File.WriteAllText(au, "Account Name,Account Number,Currency,EOD balance\n");
Assert(new BankAdapter().Process(context).Diagnostics.Count(d => d.Message.Contains("没有银行")) == 2, "Valid empty AU/GL each require explicit zero-record diagnostic");

// A fatal first row must stop before the deliberately malformed next record is scanned.
File.WriteAllText(equity, "dt_report,report_jurisdiction,is_ib,positive_equity,currency\n2026-10-25,ASIC,INVALID,1,USD\n\"unterminated");
Reject(() => new EquityAdapter().Process(context), "is_ib", "Fatal Equity row must fail fast instead of scanning and accumulating errors");

File.WriteAllText(au, "Account Name,Account Number,Currency,EOD balance\nClient,0001,USD,100\nSettle,0002,USD,50\n");
foreach (var (company, accountType, clientType) in new[]
{
    ("", "Clients", "Retail"), ("TMC", "", "Retail"), ("TMC", "Clients", ""),
    ("TMC", "Unknown", "Retail"), ("TMC", "Clients", "Unknown"), ("", "", "")
})
{
    WriteMapping(company, accountType, clientType, "NA");
    var invalid = new BankAdapter().Process(context);
    Assert(invalid.Diagnostics.Any(d => d.Severity == "Error" && d.Message.Contains("行 2") && d.Message.Contains("行 3")), "Matched invalid mapping aliases must identify both rows");
    Assert(invalid.Rows.Single(r => r["Account Number"]!.Equals("0001"))["Account Type"] == null, "Invalid mapping must not assign an arbitrary classification");
}
WriteMapping("TMC", "Clients", "Retail", "NA");
var valid = new BankAdapter().Process(context);
Assert(valid.Diagnostics.All(d => d.Severity != "Error"), "Legitimate Settle NA must remain valid");
var settle = valid.Rows.Single(r => r["Account Number"]!.Equals("0002"));
Assert(settle["Client Type"]!.Equals("NA") && settle["Settle Type"]!.Equals("NA"), "NA fields must retain their context-specific source values");

var ratesPath = Path.Combine(directory, "rates.xlsx");
using (var book = new XLWorkbook())
{
    var sheet = book.AddWorksheet("Rates");
    sheet.Cell(1, 1).Value = "Symb";
    sheet.Cell(1, 2).Value = "Units per USD";
    sheet.Cell(2, 1).Value = "USD";
    sheet.Cell(2, 2).Value = 1.25;
    sheet.Cell(2, 2).Style.NumberFormat.Format = "0.0";
    Assert(ExcelNumbers.Read(sheet.Cell(2, 2), "Synthetic amount") == 1.25m, "Shared numeric reader must ignore display rounding");
    sheet.Cell(3, 2).FormulaA1 = "5/2";
    Reject(() => ExcelNumbers.Read(sheet.Cell(3, 2), "Synthetic amount"), "缓存", "Consumed formula without numeric cache must fail");
    book.SaveAs(ratesPath);
}
Assert(DailyRates.Load(ratesPath).Convert(100m, "USD") == 80m, "Daily rate must preserve 1.25 despite display 1.3");
using (var book = new XLWorkbook(ratesPath))
{
    book.Worksheet(1).Cell(2, 2).FormulaA1 = "5/4";
    book.Save();
}
Reject(() => DailyRates.Load(ratesPath), "缓存", "Daily rates must follow strict consumed formula cache policy");
Console.WriteLine("PASS: header-only contracts, normalized header duplicates, explicit valid empty inputs, fatal Equity fail-fast, matched Bank classification, legitimate Settle NA, shared raw/cache numeric policy.");
Console.WriteLine("Synthetic evidence: " + directory);

void WriteMapping(string company, string accountType, string clientType, string settleType)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("BANK MAPPING");
    string[] headers = ["Account Name", "Account Number", "Company", "Branch", "Account Type", "Client Type", "Settle Type", "Source File"];
    for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
    string[][] rows = [
        ["Client", "0001", company, "TMA", accountType, clientType, settleType, "NAB AU"],
        ["Alias", "0001", company, "ANZO", accountType, clientType, settleType, "NAB AU.csv"],
        ["Settle", "0002", "TMC", "TMA", "Settle", "NA", "NA", "NAB AU"],
        ["Unrelated legacy", "0099", "", "OLD", "", "", "", "Other bank"]
    ];
    for (var r = 0; r < rows.Length; r++) for (var c = 0; c < headers.Length; c++) sheet.Cell(r + 2, c + 1).Value = rows[r][c];
    workbook.SaveAs(mapping);
}
static void Reject(Action action, string expected, string message)
{
    try { action(); }
    catch (InvalidDataException error) { if (error.Message.Contains(expected)) return; throw new Exception(message + ": " + error.Message); }
    throw new Exception(message);
}
static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
