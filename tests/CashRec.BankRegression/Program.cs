using CashRec;
using ClosedXML.Excel;

var directory = Directory.CreateTempSubdirectory("cashrec-bank-regression-").FullName;
var au = Path.Combine(directory, "NAB AU.csv");
var gl = Path.Combine(directory, "NAB GL.csv");
File.WriteAllText(au, "sep=,\nBalance report\nTotal balance,999999\n\nReport\n\nDetails\nAccount Name,Account Number,Currency,EOD balance\nAlias one,001-234,USD,100\nDifferent real account,001235,USD,100\nTotal balance,,,999999\n");
File.WriteAllText(gl, "sep=,\nAccount Name,Account Number,Currency,EOD balance\nSettle account,000777,USC,100\n");
var mapping = Path.Combine(directory, "mapping.xlsx");
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.AddWorksheet("BANK MAPPING");
    string[] headers = ["Account Name", "Account Number", "Company", "Branch", "Account Type", "Settle Type", "Client Type", "Bank", "Source File"];
    for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
    string[][] rows = [
        ["Alias one", "001234", "TMC", "ANZO", "Clients", "NA", "Retail", "NAB", "NAB AU.csv"],
        ["Alias two", "001234", "TMC", "TMNZ", "Clients", "NA", "Retail", "NAB", "nab au"],
        ["Second", "001235", "TMC", "TMA", "Clients", "NA", "Wholesale", "NAB", "Bank/NAB AU.csv"],
        ["Settle", "000777", "TMC", "ANZO", "Settle", "Others", "NA", "NAB", "NAB GL.csv"]
    ];
    for (var r = 0; r < rows.Length; r++) for (var c = 0; c < rows[r].Length; c++) sheet.Cell(r + 2, c + 1).Value = rows[r][c];
    workbook.SaveAs(mapping);
}
var context = new RunContext(new DateTime(2026, 10, 25), directory,
    new() { ["BankAU"] = au, ["BankGL"] = gl }, mapping, new DailyRates(new() { ["USD"] = 1m, ["USC"] = 100m }));
var table = new BankAdapter().Process(context);
Assert(table.Rows.Count == 3, "Alias mappings must not duplicate balances; real equal-amount accounts must remain");
Assert(table.Rows.Sum(row => (decimal)row["USD Equivalent"]!) == 201m, "Account EOD balances only; no total-row double counting");
Assert(table.Rows.Count(row => (string)row["Branch"]! == "TMA") == 2, "NAB AU overrides legacy Branch");
Assert(table.Rows.Single(row => (string)row["Source"]! == "NAB GL")["Branch"]!.Equals("TMG"), "NAB GL overrides legacy Branch");
Assert(table.Rows[0]["Account Number"]!.Equals("001234"), "Leading zeroes preserved and only display separators removed");
Assert(table.Diagnostics.All(d => d.Severity != "Error"), "Valid duplicate aliases must not be errors");
Assert(table.Diagnostics.Any(d => d.Severity == "Warning" && d.Message.Contains("ANZO")), "Legacy branch conflicts visible");

using (var workbook = new XLWorkbook(mapping))
{
    workbook.Worksheet("BANK MAPPING").Cell(3, 5).Value = "Settle";
    workbook.Save();
}
var conflicting = new BankAdapter().Process(context);
Assert(conflicting.Diagnostics.Any(d => d.Severity == "Error" && d.Message.Contains("分类冲突") && d.Message.Contains("2, 3")), "Conflicting classification must identify mapping rows");
Assert(conflicting.Rows.Single(row => row["Account Number"]!.Equals("001234"))["Account Type"] == null, "Conflicting classification must not guess");

File.AppendAllText(gl, "Unknown,009999,USD,15\n");
var unknown = new BankAdapter().Process(context);
Assert(unknown.Diagnostics.Any(d => d.Severity == "Error" && d.Message.Contains("009999")), "Unknown account must be visible for mapping repair");
Assert(unknown.Rows.Any(row => row["Account Number"]!.Equals("009999")), "Unknown source row must remain available with diagnostic");
Console.WriteLine("PASS: NAB front matter/totals, source priority, leading zeros, alias dedup, equal real balances, conflicts, unknown account.");
Console.WriteLine("Synthetic evidence: " + directory);
static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
