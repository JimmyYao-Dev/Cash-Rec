using CashRec;

try { Run(); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run()
{
    var report = new MonthlyReport(new DateTime(2026, 10, 25));
    var total = report.Rows["Total"];
    foreach (var (field, value) in new Dictionary<string, decimal>
    {
        ["Settle"] = 100, ["NET PSP"] = 70, ["NET BW"] = 10, ["NET WS"] = -8,
        ["Insto DIFF"] = 1, ["Retail DIFF"] = 2, ["Wholesale DIFF"] = 3, ["PSP DIFF"] = 4, ["PSP Pending DIFF"] = 5,
        ["Insto Pending DIFF"] = 1000, ["Retail Pending DIFF"] = 2000, ["Wholesale Pending DIFF"] = 3000, ["Settle DIFF"] = 10000,
    }) total[field] = value;
    MonthlyBreakdownCalculator.Apply(report);
    var row = report.Rows["Break down"];
    Check("Settle Bal", 100); Check("Client Bank Move", 15); Check("MT4", 80); Check("WS DEP", -8); Check("Diff", -57);
    Assert((DateTime)row["Date"]! == report.Date && report.BreakdownComplete, "Complete difference has the selected report date");
    Assert(((MonthlyFormula)row["Diff"]!).Expression == "[Client Bank Move]-[MT4]-[WS DEP]", "Difference retains linked subtraction without rounding");
    foreach (var field in new[] { "To LP", "Bank Fee", "Bank fee-conversion", "Time Diff", "Check", "TMA SETTLE PNL", "TMG SETTLE PNL", "To Settle WS", "Commission" })
        Assert(!row.ContainsKey(field), "This stage must not manufacture manual values or next-stage automatic fields");
    Assert(row["Error"] == null, "Error is genuinely blank");
    total["PSP Pending DIFF"] = null;
    MonthlyBreakdownCalculator.Apply(report);
    Assert(!report.BreakdownComplete && row["Client Bank Move"] == null && row["Diff"] == null,
        "A missing source clears both dependent results instead of using zero");
    Check("Settle Bal", 100); Check("MT4", 80); Check("WS DEP", -8);
    Assert(report.Diagnostics.Any(diagnostic => diagnostic.Severity == "Incomplete"), "Partial difference is explicit");
    Console.WriteLine("PASS: five-field difference, excluded obsolete Pending/Settle DIFF, signed flows, linked fields, blank manual/next-stage areas and missing-data propagation.");

    void Check(string field, decimal expected) => Assert(row[field] is MonthlyFormula value && value.CalculatedValue == expected, field + " expected " + expected);
}
static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
