using CashRec;

try { Run(); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run()
{
    var date = new DateTime(2026, 10, 25);
    var report = new MonthlyReport(date);
    report.Rows["TMA"]["SETTLE PNL"] = 11m;
    report.Rows["TMG"]["SETTLE PNL"] = -7m;
    report.Rows["TMNZ"] = new() { ["SETTLE PNL"] = 9999m };
    report.Rows["Break down"]["Diff"] = 21m;
    var table = new ReportTable("TransactionExchange");
    Add("TMA", "BANK BD", " pSp  cN  commission ", -20);
    Add("TMG", "Transaction reports", "PSP Commission", -5);
    Add("TMA", "Other prepared source", "PSP Commission", 3);
    Add("TMA", "BANK BD", "PSP Commission", 0);
    Add("TMA", "BANK BD", "Bank Fee", -8);
    Add("TTG", "BANK BD", "PSP Commission", -999);
    Add("TMA", "BANK BD", "PSP Commission", -888, date.AddDays(-1));
    var manual = new Dictionary<string, decimal?> { ["To LP"] = -5m, ["Bank Fee"] = 0m, ["Bank fee-conversion"] = null, ["Time Diff"] = -2.5m };
    MonthlySettlementCalculator.Apply(report, [table], manual);
    var row = report.Rows["Break down"];
    Check("TMA SETTLE PNL", 11); Check("TMG SETTLE PNL", -7); Check("To Settle WS", 0); Check("Commission", -22); Check("Check", 46.5m);
    Assert(row["To Settle WS"] is decimal, "J is numeric zero, not an empty cell or formula");
    Assert(((MonthlyFormula)row["TMA SETTLE PNL"]!).Expression == "[TMA:SETTLE PNL]", "H references only same-report TMA field, never TMNZ");
    Assert(((MonthlyFormula)row["Check"]!).Expression == "[Diff]-[To LP]-[TMA SETTLE PNL]-[TMG SETTLE PNL]-[To Settle WS]-[Commission]-[Bank Fee]-[Bank fee-conversion]-[Time Diff]", "Check retains every original signed subtraction without ROUND or thresholds");
    Assert(report.CheckComplete && table.Rows.Count == 7, "Nonzero Check completes automatic generation and does not discard transactions");
    report.Rows["Break down"]["Diff"] = null;
    MonthlySettlementCalculator.Apply(report, [table], manual);
    Assert(!report.CheckComplete && row["Check"] == null && report.Diagnostics.Any(diagnostic => diagnostic.Source == "Check" && diagnostic.Severity == "Incomplete"), "Unknown Diff cannot become a numeric Check");
    Check("Commission", -22); Check("TMA SETTLE PNL", 11);
    Console.WriteLine("PASS: same-day settlement fields, TMNZ exclusion, fixed J0, all-Source signed commission aliases, USD basis, manual blanks/zero/negative values and linked Check.");

    void Add(string branch, string source, string type, decimal usd, DateTime? day = null) => table.AddRow(new()
    {
        ["Branch"] = branch, ["Source"] = source, ["Transaction Type"] = type, ["Date"] = day ?? date,
        ["Source Date"] = new DateTime(1999, 1, 1), ["USD Equivalent"] = usd, ["Amount"] = 10000m, ["Currency"] = "AUD",
    });
    void Check(string field, decimal expected) => Assert(Number(row[field]) == expected, field + " expected " + expected);
}
static decimal Number(object? value) => value switch { decimal number => number, MonthlyFormula formula => formula.CalculatedValue, _ => throw new InvalidOperationException("Expected numeric automatic value") };
static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
