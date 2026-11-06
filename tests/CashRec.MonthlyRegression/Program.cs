using System.Security.Cryptography;
using System.Text.Json;
using CashRec;
using ClosedXML.Excel;

var privateMode = args.Length > 0 && args[0] == "--private-classification";
try
{
    if (privateMode) CheckPrivateClassification(args[1], args[2]);
    else Run(args[0], args.Contains("--baseline-only"));
}
catch (Exception error)
{
    Console.Error.WriteLine(privateMode ? "Private classification validation failed: " + error.GetType().Name : error);
    Environment.Exit(1);
}

static void Run(string fixture, bool baselineOnly)
{
    var workspace = Path.Combine(Path.GetTempPath(), "cashrec-monthly-" + Guid.NewGuid().ToString("N"));
    CopyDirectory(Path.GetFullPath(fixture), workspace);
    CheckRetailPriority(workspace);
    var date = new DateTime(2026, 10, 25);
    var current = Path.Combine(workspace, "2026.10.25", "Needed files");
    var prior = Path.Combine(workspace, "2026.10.22", "Needed files");
    var result = ReportTaskRunner.Run(workspace, date);
    Assert(result.MonthlyStageComplete && result.MonthlyOutput != null && !result.Complete,
        "Monthly stage must complete without claiming downstream stages: " + string.Join(";", result.Diagnostics.Select(d => d.Message)));
    using (var book = new XLWorkbook(result.MonthlyOutput!))
    {
        Assert(book.Worksheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "TMA", "TMG", "Total", "Break down" }), "Exactly four monthly sheets");
        Money(book, "TMA", date, "H", 0); Money(book, "TMA", date, "I", 880); Money(book, "TMA", date, "J", 10);
        Money(book, "TMA", date, "K", 103); Money(book, "TMA", date, "L", 100); Money(book, "TMA", date, "M", 100);
        Money(book, "TMA", date, "N", 81); Money(book, "TMG", date, "N", -7); Money(book, "Total", date, "N", 74);
        Money(book, "Total", date, "P", 10); Money(book, "Total", date, "R", 80); Money(book, "Total", date, "T", 50);
        Money(book, "Total", date, "V", 53); Money(book, "Total", date, "W", 70); Money(book, "Total", date, "X", 20);
        foreach (var name in new[] { "TMA", "TMG", "Total" })
        {
            var sheet = book.Worksheet(name); var row = DateRow(sheet, date);
            foreach (var column in new[] { "Q", "S", "U" }) Assert(sheet.Cell(column + row).IsEmpty(), "Removed Pending DIFF placeholders stay empty");
            Assert(sheet.Cell("N2").GetString() == "SETTLE PNL", "SETTLE PNL follows column deletion to N");
        }
        Money(book, "TMA", date, "B", 100); Money(book, "TMA", date, "C", -30); Money(book, "TMA", date, "D", 70);
        Money(book, "TMA", date, "E", 100); Money(book, "TMA", date, "F", -8);
        Money(book, "TMG", date, "B", 100); Money(book, "TMG", date, "C", -20); Money(book, "TMG", date, "D", 80);
        Money(book, "Total", date, "D", 150); Money(book, "Total", date, "E", 100); Money(book, "Total", date, "F", -8);
        Assert(book.Worksheet("Total").Cell("N" + DateRow(book.Worksheet("Total"), date)).HasFormula, "Total is a live two-entity field formula");
        var breakdown = book.Worksheet("Break down");
        for (var column = 2; column <= 16; column++) Assert(breakdown.Cell(2, column).IsEmpty(), "#9 Breakdown values remain genuinely absent");
        Assert(breakdown.Cell("G2").Style.NumberFormat.Format == "\"$\"#,##0.00;[Red]\\-\"$\"#,##0.00", "Original To LP format retained");
    }
    using (var json = JsonDocument.Parse(File.ReadAllText(result.DiagnosticPath)))
    {
        Assert(json.RootElement.GetProperty("MonthlyStage").GetString() == "Flows"
            && json.RootElement.GetProperty("MonthlyStageComplete").GetBoolean()
            && !json.RootElement.GetProperty("Complete").GetBoolean(), "Task diagnostic distinguishes complete stage from incomplete full monthly task");
    }
    var transactions = BalanceWorkbookReader.ReadRows(result.DailyOutput!, "TransactionExchange", ["Source", "Company", "Organization", "Transaction Type"]);
    var bankRows = transactions.Where(row => (string?)row["Source"] == "BANK BD").ToList();
    Assert(bankRows.Any(row => (string?)row["Company"] == "TMA-Retail" && (string?)row["Organization"] == "TMC"), "Source Company retained while organization normalized");
    foreach (var category in new[] { "Posting TF TD", "MISC CIMB TF TD", "CFD TF TD", "Rebate TF TD", "Pending TF TD" })
        Assert(bankRows.Any(row => (string?)row["Transaction Type"] == category), "Verified settlement category must not be lost to Other: " + category);
    if (baselineOnly)
    {
        Console.WriteLine("PASS: synthetic monthly baseline; independent reader may inspect this developer-run artifact.");
        Console.WriteLine("Synthetic developer workspace: " + workspace);
        return;
    }

    // The previous set includes an account that disappeared; the current set includes a new key.
    var au = Path.Combine(current, "Bank", "NAB AU.csv");
    var mapping = Path.Combine(workspace, "db", "tmc", "account_mapping.xlsx");
    var originalAu = File.ReadAllBytes(au); var originalMapping = File.ReadAllBytes(mapping);
    File.WriteAllLines(au, File.ReadAllLines(au).Where(line => !line.Contains("001-1003")));
    File.AppendAllText(au, "New synthetic Wholesale,001-1999,USD,0,0,0,0,20\n");
    using (var book = new XLWorkbook(mapping))
    {
        var sheet = book.Worksheet("BANK MAPPING"); var next = sheet.LastRowUsed()!.RowNumber() + 1;
        var values = new Dictionary<string, string> { ["Account Name"] = "New synthetic Wholesale", ["Account Number"] = "001-1999", ["Company"] = "TMC", ["Branch"] = "TMA", ["Account Type"] = "Clients", ["Settle Type"] = "NA", ["Client Type"] = "Wholesale", ["Bank"] = "NAB", ["Source File"] = "NAB AU" };
        foreach (var header in sheet.Row(1).CellsUsed()) if (values.TryGetValue(header.GetString(), out var value)) sheet.Cell(next, header.Address.ColumnNumber).Value = value;
        book.Save();
    }
    result = ReportTaskRunner.Run(workspace, date);
    Assert(result.MonthlyStageComplete, "Complete snapshot group comparison is valid even when a new daily row has no matched prior key");
    using (var book = new XLWorkbook(result.MonthlyOutput!)) Money(book, "TMA", date, "T", 10);
    File.WriteAllBytes(au, originalAu); File.WriteAllBytes(mapping, originalMapping);

    // The exact previous Pending snapshot is mandatory, even when daily balances exist.
    ReportRunner.Run(workspace, new DateTime(2026, 10, 22));
    var priorBankBd = Path.Combine(prior, "PSP+BANK", "BANK BD.xlsx");
    File.Move(priorBankBd, priorBankBd + ".held");
    try
    {
        result = ReportTaskRunner.Run(workspace, date);
        Assert(!result.MonthlyStageComplete && result.MonthlyOutput != null, "Missing prior Pending is partial, not zero or a fallback to today's snapshot");
        using var book = new XLWorkbook(result.MonthlyOutput!);
        var sheet = book.Worksheet("Total"); Assert(sheet.Cell("W" + DateRow(sheet, date)).IsEmpty(), "Missing previous Pending clears stale DIFF");
        Money(book, "Total", date, "R", 80);
    }
    finally { File.Move(priorBankBd + ".held", priorBankBd); }

    // A historical run reads its own unchanged file, never a newer cumulative snapshot.
    var bankBd = Path.Combine(current, "PSP+BANK", "BANK BD.xlsx");
    using (var book = new XLWorkbook(bankBd))
    {
        var sheet = book.Worksheet("PSP On The Way"); var row = sheet.RowsUsed().Single(row => row.Cell(17).GetString() == "TMA");
        row.Cell(18).Value = 2000; row.Cell(19).Value = 1001; row.Cell(20).Value = 999;
        book.Save();
    }
    var historical = ReportTaskRunner.Run(workspace, new DateTime(2026, 10, 22));
    Assert(historical.MonthlyStageComplete, "Historical fixture has its own prior date");
    using (var book = new XLWorkbook(historical.MonthlyOutput!)) Money(book, "TMA", new DateTime(2026, 10, 22), "L", 80);

    // A consumed source formula without a cache prevents monthly replacement.
    var previousHash = Hash(historical.MonthlyOutput!);
    using (var book = new XLWorkbook(bankBd))
    {
        var sheet = book.Worksheet("PSP On The Way"); var row = sheet.RowsUsed().Single(row => row.Cell(17).GetString() == "TMA");
        row.Cell(20).FormulaA1 = "R" + row.RowNumber() + "-S" + row.RowNumber(); book.Save();
    }
    result = ReportTaskRunner.Run(workspace, date);
    Assert(!result.MonthlyStageComplete && result.MonthlyOutput == null && result.Diagnostics.Any(d => d.Message.Contains("缓存")), "Missing cached Pending is an explicit source error");
    Assert(Hash(historical.MonthlyOutput!) == previousHash, "Rejected Pending must not overwrite existing monthly result");
    Console.WriteLine("PASS: monthly balances/complete-set DIFF, original layout/formats, signed settlement categories, stage diagnostics, historical snapshots, missing previous Pending and cache errors.");
    Console.WriteLine("Synthetic developer workspace: " + workspace);
}

static void CopyDirectory(string source, string target)
{
    Directory.CreateDirectory(target);
    foreach (var path in Directory.EnumerateFiles(source)) File.Copy(path, Path.Combine(target, Path.GetFileName(path)));
    foreach (var directory in Directory.EnumerateDirectories(source)) CopyDirectory(directory, Path.Combine(target, Path.GetFileName(directory)));
}
static int DateRow(IXLWorksheet sheet, DateTime date) => sheet.RowsUsed().Single(row => row.RowNumber() > MonthlyLayout.HeaderRows(sheet.Name)
    && row.Cell(1).TryGetValue<DateTime>(out var value) && value.Date == date.Date).RowNumber();
static void Money(XLWorkbook book, string sheet, DateTime date, string column, decimal expected)
{
    var cell = book.Worksheet(sheet).Cell(column + DateRow(book.Worksheet(sheet), date));
    Assert(ExcelNumbers.Read(cell, sheet + " " + column) == expected, sheet + " " + column + " must equal " + expected);
}
static string Hash(string path) { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)); }
static void Assert(bool value, string label) { if (!value) throw new InvalidOperationException(label); }

static void CheckRetailPriority(string workspace)
{
    var path = Path.Combine(workspace, "synthetic-retail-priority.xlsx");
    (string Comment, string Rational, string Expected, string Company)[] cases =
    [
        ("settlement", "settle transfer", "Settle", "TMA-Retail"),
        ("posting CFD", "settle + TF", "CFD TF TD", "TMA-Retail"),
        ("cfd rebate PENDING", "settle + tf", "Pending TF TD", "TMA-Retail"),
        ("settle cfd", "settle + tf + fee", "Bank Fee", "TMA-Retail"),
        ("pending", "TF + interest + PSP SETTLE", "PSP Settle In", "TMA-Retail"),
        ("ordinary", "TF + SURCHARGE", "Bank TF", "TMA-Retail"),
        ("ordinary", "TF + deposit", "BW Dep", "TMA-Retail"),
        ("ordinary", "TF + deposit + WITHdrawal", "BW WD", "TMA-Retail"),
        ("ordinary", "TF + psp + dep", "Bank TF", "TMA-Retail"),
        ("ordinary", "TF + insto + with", "Other", "TMA-Retail"),
        ("ordinary", "unrelated adjustment", "Other", "TMA-Retail"),
        ("ordinary", "Adjustment", "Adjustment", "TMA-Retail"),
        ("PSP Commission", "PSP Commission", "PSP Commission", "TMA-Retail"),
        ("settle cfd", "settle + tf + fee", "Other", "TMA-Wholesale"),
    ];
    using (var book = new XLWorkbook())
    {
        var sheet = book.AddWorksheet("BANK BD");
        string[] headers = ["Credit", "Debit", "Comment", "Rational", "Currency", "Company", "Amount"];
        for (var column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).Value = headers[column];
        for (var i = 0; i < cases.Length; i++)
        {
            sheet.Cell(i + 2, 3).Value = cases[i].Comment;
            sheet.Cell(i + 2, 4).Value = cases[i].Rational;
            sheet.Cell(i + 2, 5).Value = "USD";
            sheet.Cell(i + 2, 6).Value = cases[i].Company;
            sheet.Cell(i + 2, 7).Value = -7;
        }
        book.SaveAs(path);
    }
    var context = new RunContext(new DateTime(2026, 10, 25), workspace, new() { ["BankBD"] = path }, "", new DailyRates(new() { ["USD"] = 1m }));
    var table = new BankBdTransactionAdapter().Process(context);
    Assert(table.Diagnostics.All(diagnostic => diagnostic.Severity != "Error") && table.Rows.Count == cases.Length, "Synthetic classification inputs retained");
    for (var i = 0; i < cases.Length; i++)
    {
        var row = table.Rows.Single(row => (int)row["Source Row"]! == i + 2);
        Assert((string)row["Transaction Type"]! == cases[i].Expected, "Retail last-match priority case " + i);
        Assert((string)row["Company"]! == cases[i].Company && (string)row["Organization"]! == "TMC", "Raw company and normalized organization stay separate");
    }
}

static void CheckPrivateClassification(string workspace, string vectorsPath)
{
    var date = new DateTime(2026, 10, 25);
    var prepared = WorkspaceInputs.Inspect(workspace, date);
    using var expected = JsonDocument.Parse(File.ReadAllText(vectorsPath));
    var contract = expected.RootElement;
    Assert(Hash(prepared.Inputs["BankBD"]).Equals(contract.GetProperty("source_sha256").GetString(), StringComparison.OrdinalIgnoreCase), "Private source SHA must match the independently bound vectors");
    var context = new RunContext(date, prepared.DateDirectory, prepared.Inputs, prepared.MappingPath,
        DailyRates.Load(prepared.Inputs["Rates"], date));
    var table = new BankBdTransactionAdapter().Process(context);
    var actual = table.Rows.Where(row => string.Equals(row.GetValueOrDefault("Type")?.ToString(), "Retail", StringComparison.OrdinalIgnoreCase))
        .ToDictionary(row => (int)row["Source Row"]!);
    var matched = 0;
    foreach (var vector in contract.GetProperty("vectors").EnumerateArray())
    {
        if (actual.TryGetValue(vector.GetProperty("source_row").GetInt32(), out var row)
            && string.Equals(row["Branch"]?.ToString(), vector.GetProperty("expected_branch").GetString(), StringComparison.OrdinalIgnoreCase)
            && DataValue.Normalize(row["Transaction Type"]!.ToString()!) == DataValue.Normalize(vector.GetProperty("expected_transaction_type").GetString()!)) matched++;
    }
    var count = contract.GetProperty("expected_count").GetInt32();
    var errors = table.Diagnostics.Count(diagnostic => diagnostic.Severity == "Error");
    var report = new { Passed = actual.Count == count && matched == count && errors == 0, ExpectedCount = count, ActualRetailCount = actual.Count, MatchedCount = matched, ErrorCount = errors, Entry = "Private C# adapter developer comparison; not installed-GUI acceptance" };
    var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
    Console.WriteLine(json);
    File.WriteAllText(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(vectorsPath))!, "csharp-retail-classification-validation-private.json"), json);
    Assert(report.Passed, "Private comparison counts must agree");
}
