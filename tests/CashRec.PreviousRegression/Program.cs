using CashRec;
using ClosedXML.Excel;
using System.Text.Json;

var workspace = Path.GetFullPath(args[0]);
var date = new DateTime(2026, 10, 25);
var result = ReportRunner.Run(workspace, date);
Assert(result.IsComplete && result.OutputPath != null, string.Join(";", result.Diagnostics.Select(d => d.Message)));
CheckTotals(result.OutputPath!);
Assert(BalanceWorkbookReader.HasInitializedDep(result.OutputPath!), "DEP initialization marker");
using (var workbook = new XLWorkbook(result.OutputPath!))
{
    var sheet = workbook.Worksheet("Equity");
    var headers = sheet.Row(1).CellsUsed().ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);
    var row = sheet.RowsUsed().Skip(1).Single(r => r.Cell(headers["Branch"]).GetString() == "TMA" && r.Cell(headers["Retail Type"]).GetString() == "Standard" && r.Cell(headers["Currency"]).GetString() == "USD");
    row.Cell(headers["DEP"]).Value = 30;
    workbook.Save();
}
result = ReportRunner.Run(workspace, date);
Assert(result.IsComplete, "Rerun complete");
var equity = BalanceWorkbookReader.ReadRows(result.OutputPath!, "Equity").Single(r => (string)r["Branch"]! == "TMA" && (string)r["Retail Type"]! == "Standard" && (string)r["Currency"]! == "USD");
Assert((decimal)equity["DEP"]! == 30m && (decimal)equity["Movement"]! == 8m && (decimal)equity["PNL"]! == -22m, "Manual DEP preserved and engine PNL updated");
Assert(Directory.GetFiles(Path.Combine(workspace, "2026.10.25", "backups"), "*.xlsx").Length > 0, "Existing results backed up");

var prior = ReportRunner.Run(workspace, new DateTime(2026, 10, 22));
Assert(!prior.IsComplete && prior.OutputPath != null, "First snapshot has no fabricated earlier basis");
var priorInput = Path.Combine(workspace, "2026.10.22", "Needed files", "MT4", "Risk", "Accounting Equity Report.csv");
var held = priorInput + ".test-held";
File.Move(priorInput, held);
try { result = ReportRunner.Run(workspace, date); Assert(result.IsComplete, "Same-rules saved snapshot accepted without original CSV"); CheckTotals(result.OutputPath!); }
finally { File.Move(held, priorInput); }
var absent = ReportRunner.Run(Path.GetFullPath(args[1]), date);
Assert(!absent.IsComplete && absent.OutputPath != null && absent.Diagnostics.Any(d => d.Severity == "Incomplete"), "Missing prior explicitly incomplete");
foreach (var row in BalanceWorkbookReader.ReadRows(absent.OutputPath!, "Equity"))
    Assert(row["Prev USD Equivalent"] == null && row["Movement"] == null && row["PNL"] == null && (decimal)row["DEP"]! == 0m, "Missing previous not zero");
Console.WriteLine("Previous regression passed: same-rule raw reconstruction, nonconsecutive dates, persisted snapshots, manual DEP and explicit missing basis.");

void CheckTotals(string path)
{
    using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(workspace, "expected.json")));
    foreach (var name in new[] { "Equity", "Bank", "PSP" })
    {
        var contract = expected.RootElement.GetProperty("sheets").GetProperty(name);
        var rows = BalanceWorkbookReader.ReadRows(path, name);
        foreach (var (field, total) in new[] { ("USD Equivalent", "total_usd"), ("Prev USD Equivalent", "total_previous_usd"), ("Movement", "total_movement") })
            Assert(rows.Sum(row => (decimal)row[field]!) == contract.GetProperty(total).GetDecimal(), name + " " + field);
    }
}
static void Assert(bool value, string label) { if (!value) throw new Exception(label); }
