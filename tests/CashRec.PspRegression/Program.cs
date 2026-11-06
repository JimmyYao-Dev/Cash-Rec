using System.Text.Json;
using CashRec;

var workspace = Path.GetFullPath(args[0]);
var date = new DateTime(2026, 10, 25);
var check = WorkspaceInputs.Inspect(workspace, date);
var context = new RunContext(date, check.DateDirectory, check.Inputs, check.MappingPath, DailyRates.Load(check.Inputs["Rates"], date));
var result = new PspAdapter().Process(context);
using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(workspace, "expected.json")));
var expected = oracle.RootElement.GetProperty("sheets").GetProperty("PSP");
Assert(result.Rows.Count == expected.GetProperty("rows").GetArrayLength(), "PSP row count");
foreach (var target in expected.GetProperty("rows").EnumerateArray())
{
    var row = result.Rows.Single(row => new[] { "Source", "PSP", "Currency" }.All(key => (string)row[key]! == target.GetProperty(key).GetString()));
    Assert((string)row["Branch"]! == target.GetProperty("Branch").GetString(), "Source entity");
    foreach (var key in new[] { "Amount", "USD Equivalent" }) Assert((decimal)row[key]! == target.GetProperty(key).GetDecimal(), key);
    Assert(row["Prev USD Equivalent"] == null && row["Movement"] == null, "No invented previous");
}
Assert(result.Rows.Sum(r => (decimal)r["USD Equivalent"]!) == 253m, "Independent USD total");
Assert(result.Rows.Any(r => (string)r["PSP"]! == "NewGateway" && (string)r["Mapping Status"]! == "待补映射"), "Unmapped PSP retained");
if (args.Length > 1)
{
    var missing = WorkspaceInputs.Inspect(Path.GetFullPath(args[1]), date);
    var failed = false;
    try
    {
        var rejected = new PspAdapter().Process(new(date, missing.DateDirectory, missing.Inputs, missing.MappingPath, DailyRates.Load(missing.Inputs["Rates"], date)));
        failed = rejected.Rows.Count == 0 && rejected.Diagnostics.Any(diagnostic => diagnostic.Severity == "Error" && diagnostic.Code == DiagnosticCode.MissingPspDate);
    }
    catch (InvalidDataException error) { failed = error.Message.Contains("2026.10.25"); }
    Assert(failed, "Missing PSP date rejected");
}
Console.WriteLine("PSP regression passed: exact D, source scope, mappings, raw amounts, currency units and no fabricated previous.");
static void Assert(bool condition, string label) { if (!condition) throw new Exception(label); }
