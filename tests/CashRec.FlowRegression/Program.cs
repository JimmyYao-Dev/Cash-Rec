using CashRec;

try { Run(); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run()
{
    var date = new DateTime(2026, 10, 25);
    var table = new ReportTable("TransactionExchange");
    Add("TMA", "Transaction reports", "PSP", "PSP Dep", "", 100);
    Add("TMA", "Transaction reports", "PSP", "PSP Dep", "", 100);
    Add("TMA", "Transaction reports", "PSP", "PSP WITH", "", -30);
    Add("TMA", "Transaction reports", "PSP", "PSP WITH", "", 0);
    Add("TMA", "BANK BD", "PSP", "PSP Dep", "", 999);
    Add("TMA", "Transaction reports", "Retail", "BW Dep", "MT4 BW", 100);
    Add("TMA", "Transaction reports", "Retail", "BW WD", "MT4 BW", -10);
    Add("TMA", "Transaction reports", "Retail", "Bank Fee", "", 999);
    Add("TMA", "BANK BD", "Retail", "BW Dep", "MT4 BW", 777);
    Add("TMA", "BANK BD", "Wholesale", "Withdrawal", "", -8);
    Add("TMA", "BANK BD", "Wholesale", "Deposit", "", 8);
    Add("TMA", "BANK BD", "Wholesale", "Other", "", 0);
    Add("TMA", "Transaction reports", "Wholesale", "Deposit", "", 99);
    Add("TMA", "BANK BD archive", "Wholesale", "Deposit", "", 9999);
    Add("TTG", "BANK BD", "Wholesale", "Deposit", "", 9000);
    Add("TMA", "Transaction reports", "PSP", "Other", "", null);
    Add("TMA", "Transaction reports", "PSP", "PSP Dep", "", 5000, date.AddDays(-1));
    Add("TMA", " transaction REPORTS ", " psp ", " psp   dep ", "", 5);
    Add("TMG", "Transaction reports", "PSP", "PSP Dep", "", 40);
    Add("TMG", "Transaction reports", "PSP", "PSP WITH", "", -5);
    Add("TMG", "Transaction reports", "Retail", "BW Dep", "MT4 BW", 0);
    Add("TMG", "BANK BD", "Wholesale", "Deposit", "", 12);
    Add("TMG", "BANK BD", "Wholesale", "Withdrawal", "", -2);
    var report = new MonthlyReport(date);
    MonthlyFlowCalculator.Apply(report, [table]);
    Check("TMA", "PSP DEP", 205); Check("TMA", "PSP WITH", -30); Check("TMA", "NET PSP", 175);
    Check("TMA", "NET BW", 90); Check("TMA", "NET WS", 0);
    Check("TMG", "PSP DEP", 40); Check("TMG", "PSP WITH", -5); Check("TMG", "NET PSP", 35);
    Check("TMG", "NET BW", 0); Check("TMG", "NET WS", 10);
    Check("Total", "PSP DEP", 245); Check("Total", "PSP WITH", -35); Check("Total", "NET PSP", 210);
    Check("Total", "NET BW", 90); Check("Total", "NET WS", 10);
    Assert(((MonthlyFormula)report.Rows["TMA"]["NET PSP"]!).Expression == "[PSP DEP]+[PSP WITH]", "Signed DEP and WITH remain a linked addition formula");
    Assert(report.Rows["Total"].Values.All(value => value is MonthlyFormula), "Every consolidated flow remains a live entity-sum formula");
    Assert(report.FlowsComplete && !report.StageComplete, "Flow calculation must not hide missing balance prerequisites");
    Assert(table.Rows.Count == 23 && table.Rows.Count(row => row["USD Equivalent"] is decimal number && number == 0) == 3,
        "Zero and same-amount source records are retained");

    var zero = new MonthlyReport(date);
    MonthlyFlowCalculator.Apply(zero, [new ReportTable("TransactionExchange")]);
    foreach (var field in MonthlyFlowCalculator.Fields)
        Assert(Number(zero.Rows["TMA"][field]) == 0 && Number(zero.Rows["Total"][field]) == 0, "Valid empty flow groups are numeric zero");
    var invalid = new ReportTable("TransactionExchange");
    invalid.AddRow(new() { ["Date"] = date, ["Branch"] = "TMA", ["Source"] = "Transaction reports", ["Type"] = "PSP", ["Transaction Type"] = "PSP Dep", ["USD Equivalent"] = null });
    Reject(() => MonthlyFlowCalculator.Apply(new MonthlyReport(date), [invalid]), "数值", "A consumed missing USD value cannot become zero");
    invalid.Rows[0]["Date"] = null;
    Reject(() => MonthlyFlowCalculator.Apply(new MonthlyReport(date), [invalid]), "日期", "Missing report date must be reported");
    Console.WriteLine("PASS: signed flow sums, strict source/type/date isolation, same-amount records, zeros, USD basis, empty groups, formula values and consumed-data errors.");

    void Add(string branch, string source, string type, string transactionType, string exchangeType, decimal? usd, DateTime? reportDate = null)
    {
        table.AddRow(new()
        {
            ["Date"] = reportDate ?? date, ["Source Date"] = new DateTime(1999, 1, 1),
            ["Branch"] = branch, ["Source"] = source, ["Type"] = type,
            ["Transaction Type"] = transactionType, ["Exchange Type"] = exchangeType,
            ["USD Equivalent"] = usd, ["Amount"] = 40000m, ["Currency"] = "AUD",
            ["Transaction ID"] = "SYN-" + table.Rows.Count,
        });
    }
    void Check(string branch, string field, decimal expected) => Assert(Number(report.Rows[branch][field]) == expected, branch + " " + field + " must equal " + expected);
}
static decimal Number(object? value) => value switch { decimal number => number, MonthlyFormula formula => formula.CalculatedValue, _ => throw new InvalidOperationException("Expected a computed flow value") };
static void Reject(Action action, string expected, string message)
{
    try { action(); }
    catch (InvalidDataException error) { if (error.Message.Contains(expected)) return; throw; }
    throw new InvalidOperationException(message);
}
static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
