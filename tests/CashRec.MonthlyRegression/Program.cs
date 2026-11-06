using System.Security.Cryptography;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
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
    Assert(result.MonthlyStageComplete && result.MonthlyOutput != null && result.Complete,
        "Single-day automatic reports must complete without requiring Check to be zero: " + string.Join(";", result.Diagnostics.Select(d => d.Message)));
    using (var book = new XLWorkbook(result.MonthlyOutput!))
    {
        Assert(book.Worksheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "TMA", "TMG", "Total", "Break down" }), "Exactly four monthly sheets");
        Money(book, "TMA", date, "H", 0); Money(book, "TMA", date, "I", 880); Money(book, "TMA", date, "J", 10);
        Money(book, "TMA", date, "K", 103); Money(book, "TMA", date, "L", 100); Money(book, "TMA", date, "M", 100);
        Money(book, "TMA", date, "N", 385); Money(book, "TMG", date, "N", -7); Money(book, "Total", date, "N", 378);
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
        Money(book, "Break down", date, "B", 100); Money(book, "Break down", date, "C", 263);
        Money(book, "Break down", date, "D", 250); Money(book, "Break down", date, "E", -8); Money(book, "Break down", date, "F", 21);
        Money(book, "Break down", date, "H", 385); Money(book, "Break down", date, "I", -7);
        Money(book, "Break down", date, "J", 0); Money(book, "Break down", date, "K", -25); Money(book, "Break down", date, "P", -332);
        foreach (var column in new[] { 7, 12, 13, 14, 15 })
            Assert(breakdown.Cell(2, column).IsEmpty(), "Manual areas and Error remain genuinely absent");
        Assert(breakdown.Cell("P2").FormulaA1 == "F2-G2-H2-I2-J2-K2-L2-M2-N2", "Original Check subtraction formula remains unchanged");
        Assert(breakdown.Cell("G2").Style.NumberFormat.Format == "\"$\"#,##0.00;[Red]\\-\"$\"#,##0.00", "Original To LP format retained");
    }
    using (var json = JsonDocument.Parse(File.ReadAllText(result.DiagnosticPath)))
    {
        Assert(json.RootElement.GetProperty("MonthlyStage").GetString() == "Check"
            && json.RootElement.GetProperty("MonthlyStageComplete").GetBoolean()
            && json.RootElement.GetProperty("Complete").GetBoolean(), "Task completion means this automatic single-day task, not manual reconciliation");
    }
    var transactions = BalanceWorkbookReader.ReadRows(result.DailyOutput!, "TransactionExchange", ["Source", "Company", "Organization", "Transaction Type"]);
    var bankRows = transactions.Where(row => (string?)row["Source"] == "BANK BD").ToList();
    Assert(bankRows.Any(row => (string?)row["Company"] == "TMA-Retail" && (string?)row["Organization"] == "TMC"), "Source Company retained while organization normalized");
    foreach (var category in new[] { "Posting TF TD", "MISC CIMB TF TD", "CFD TF TD", "Rebate TF TD", "Pending TF TD" })
        Assert(bankRows.Any(row => (string?)row["Transaction Type"] == category), "Verified settlement category must not be lost to Other: " + category);
    Assert(BalanceWorkbookReader.HasCurrentClassification(result.DailyOutput!), "Daily output declares the current classification policy");
    Assert(BalanceWorkbookReader.HasCurrentClassification(result.MonthlyOutput!)
        && BalanceWorkbookReader.HasCurrentClassification(result.MonthlyOutput!, date), "New monthly output has whole-file and per-date classification coverage");
    if (baselineOnly)
    {
        Console.WriteLine("PASS: synthetic monthly baseline; independent reader may inspect this developer-run artifact.");
        Console.WriteLine("Synthetic developer workspace: " + workspace);
        return;
    }
    CheckClassificationHistory(workspace, date, result.MonthlyOutput!, result.DailyOutput!);

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
        var breakdown = book.Worksheet("Break down"); var breakdownRow = DateRow(breakdown, date);
        Assert(breakdown.Cell(breakdownRow, 3).IsEmpty() && breakdown.Cell(breakdownRow, 6).IsEmpty() && breakdown.Cell(breakdownRow, 16).IsEmpty(), "Missing source DIFF clears stale Client Bank Move, Diff and Check");
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
        ("CFD TF TD", "memo", "CFD TF TD", "TMA-Retail"),
        ("MISC CIMB TF TD", "memo", "MISC CIMB TF TD", "TMA-Retail"),
        ("posting", "memo", "Posting TF TD", "TMA-Retail"),
        ("Rebate", "memo", "Rebate TF TD", "TMA-Retail"),
        ("pending", "memo", "Pending TF TD", "TMA-Retail"),
        ("SETTLE", "memo", "Settle", "TMA-Retail"),
        ("cfd rebate PENDING", "memo", "Pending TF TD", "TMA-Retail"),
        ("CFD TF TD", "fee", "Bank Fee", "TMA-Retail"),
        ("SETTLE", "psp settle", "PSP Settle In", "TMA-Retail"),
        ("SETTLE", "psp credit", "Other", "TMA-Retail"),
        ("CFD TF TD", "deposit", "BW Dep", "TMA-Retail"),
        ("miscellaneous decoy", "memo", "MISC CIMB TF TD", "TMA-Retail"),
        ("MISC CIMB TF TD", "memo", "Other", "TMA-Wholesale"),
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
    var identityCoverage = true;
    var confirmedChangesOnly = true;
    var confirmedTypes = new[] { "Settle", "Posting TF TD", "MISC CIMB TF TD", "CFD TF TD", "Rebate TF TD", "Pending TF TD" }
        .Select(DataValue.Normalize).ToHashSet();
    foreach (var vector in contract.GetProperty("vectors").EnumerateArray())
    {
        if (!actual.TryGetValue(vector.GetProperty("source_row").GetInt32(), out var row)
            || !string.Equals(row["Branch"]?.ToString(), vector.GetProperty("expected_branch").GetString(), StringComparison.OrdinalIgnoreCase))
        { identityCoverage = false; continue; }
        var type = DataValue.Normalize(row["Transaction Type"]!.ToString()!);
        if (type == DataValue.Normalize(vector.GetProperty("expected_transaction_type").GetString()!)) matched++;
        else if (!confirmedTypes.Contains(type)) confirmedChangesOnly = false;
    }
    var count = contract.GetProperty("expected_count").GetInt32();
    var errors = table.Diagnostics.Count(diagnostic => diagnostic.Severity == "Error");
    var report = new { Passed = actual.Count == count && identityCoverage && confirmedChangesOnly && errors == 0,
        Policy = ReportPolicy.Classification, ExpectedCount = count, ActualRetailCount = actual.Count, UnchangedCount = matched,
        ChangedUnderConfirmedPolicy = count - matched, IdentityCoverage = identityCoverage, ConfirmedChangesOnly = confirmedChangesOnly,
        ErrorCount = errors, Entry = "Private C# adapter comparison to legacy classifications; intended keyword differences are allowed; not installed-GUI acceptance" };
    var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
    Console.WriteLine(json);
    File.WriteAllText(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(vectorsPath))!, "csharp-retail-classification-comparison-comment-keywords-private.json"), json);
    Assert(report.Passed, "Private source identities must agree and category differences must follow the confirmed policy");
}

static void CheckClassificationHistory(string workspace, DateTime currentDate, string monthlyPath, string dailyPath)
{
    var earlier = new DateTime(2026, 10, 22);
    var previousRun = ReportTaskRunner.Run(workspace, earlier);
    Assert(previousRun.MonthlyStageComplete, "Historical synthetic fixture can be generated under the new policy");
    string manualFormat;
    using (var book = new XLWorkbook(monthlyPath))
    {
        var sheet = book.Worksheet("Break down"); var oldRow = DateRow(sheet, earlier); var currentRow = DateRow(sheet, currentDate);
        sheet.Cell(oldRow, 7).Value = 0;
        sheet.Cell(oldRow, 12).Value = -12.5;
        sheet.Cell(oldRow, 13).Clear(XLClearOptions.Contents);
        sheet.Cell(oldRow, 14).Value = -3;
        sheet.Cell(currentRow, 7).Value = -9;
        sheet.Cell(currentRow, 15).Value = 17;
        manualFormat = sheet.Cell(oldRow, 12).Style.NumberFormat.Format;
        book.Worksheet("TMA").Cell("N" + DateRow(book.Worksheet("TMA"), earlier)).Value = 9999;
        book.Worksheet("TMA").Cell("N" + DateRow(book.Worksheet("TMA"), currentDate)).Value = 9999;
        var total = book.Worksheet("Total");
        foreach (var column in new[] { "Q", "S", "U" }) total.Cell(column + DateRow(total, currentDate)).Value = 10000;
        book.Save();
    }
    // Simulate the actual pre-change monthly format: no classification markers at all.
    using (var archive = ZipFile.Open(monthlyPath, ZipArchiveMode.Update))
    {
        var entry = archive.GetEntry("docProps/custom.xml")!;
        XDocument xml;
        using (var stream = entry.Open()) xml = XDocument.Load(stream);
        foreach (var property in xml.Descendants().Where(element => element.Name.LocalName == "property"
            && ((string?)element.Attribute("name") ?? "").StartsWith("CashRecClassification", StringComparison.Ordinal)).ToList()) property.Remove();
        entry.Delete();
        using var output = archive.CreateEntry("docProps/custom.xml").Open(); xml.Save(output);
    }
    using (var book = new XLWorkbook(dailyPath))
    {
        book.CustomProperties.Single(property => property.Name == "CashRecClassification").Value = "Legacy";
        var sheet = book.Worksheet("Equity");
        var headers = sheet.Row(1).CellsUsed().ToDictionary(cell => cell.GetString(), cell => cell.Address.ColumnNumber);
        var row = sheet.RowsUsed().Skip(1).Single(row => row.Cell(headers["Branch"]).GetString() == "TMA"
            && row.Cell(headers["Retail Type"]).GetString() == "Standard" && row.Cell(headers["Currency"]).GetString() == "USD");
        row.Cell(headers["DEP"]).Value = 30; book.Save();
    }
    Assert(BalanceWorkbookReader.HasCurrentRules(dailyPath) && BalanceWorkbookReader.HasCompatibleManualKeys(dailyPath)
        && !BalanceWorkbookReader.HasCurrentClassification(dailyPath), "Classification change must not invalidate balance valuation or manual keys");
    var result = ReportTaskRunner.Run(workspace, currentDate);
    Assert(result.MonthlyStageComplete && result.Diagnostics.Any(diagnostic => diagnostic.Source == "历史交易分类" && diagnostic.Severity == "Warning"), "Old history remains visible as needing recomputation");
    Assert(!BalanceWorkbookReader.HasCurrentClassification(monthlyPath)
        && BalanceWorkbookReader.HasCurrentClassification(monthlyPath, currentDate)
        && !BalanceWorkbookReader.HasCurrentClassification(monthlyPath, earlier), "Only the recomputed date gets the current policy; whole-file marker is Mixed");
    using (var book = new XLWorkbook(monthlyPath))
    {
        Assert(book.CustomProperties.Single(property => property.Name == "CashRecClassification").Value.ToString() == ReportPolicy.MixedClassification, "Unrecomputed history is not silently relabeled");
        Money(book, "TMA", currentDate, "N", 385); Money(book, "TMA", earlier, "N", 9999);
        Money(book, "Break down", currentDate, "P", -323);
        Assert(book.Worksheet("Break down").Cell(DateRow(book.Worksheet("Break down"), currentDate), 15).IsEmpty(), "Error is cleared automatically without treating it as manual input");
        var total = book.Worksheet("Total");
        foreach (var column in new[] { "Q", "S", "U" }) Assert(total.Cell(column + DateRow(total, currentDate)).IsEmpty(), "Obsolete Pending difference placeholders cannot retain old automatic values");
        CheckManual(book);
    }
    var equity = BalanceWorkbookReader.ReadRows(dailyPath, "Equity", ["Branch", "Retail Type", "Currency", "DEP"])
        .Single(row => (string?)row["Branch"] == "TMA" && (string?)row["Retail Type"] == "Standard" && (string?)row["Currency"] == "USD");
    Assert((decimal)equity["DEP"]! == 30m, "A stale classification tag must not discard manual DEP");
    result = ReportTaskRunner.Run(workspace, earlier);
    Assert(result.MonthlyStageComplete && BalanceWorkbookReader.HasCurrentClassification(monthlyPath)
        && BalanceWorkbookReader.HasCurrentClassification(monthlyPath, earlier), "Recomputing the last legacy date resolves Mixed coverage");
    using (var book = new XLWorkbook(monthlyPath))
    {
        Money(book, "TMA", earlier, "N", 325);
        CheckManual(book);
    }

    void CheckManual(XLWorkbook book)
    {
        var sheet = book.Worksheet("Break down"); var oldRow = DateRow(sheet, earlier); var currentRow = DateRow(sheet, currentDate);
        Assert(sheet.Cell(oldRow, 7).GetValue<decimal>() == 0m && sheet.Cell(oldRow, 12).GetValue<decimal>() == -12.5m
            && sheet.Cell(oldRow, 13).IsEmpty() && sheet.Cell(oldRow, 14).GetValue<decimal>() == -3m
            && sheet.Cell(currentRow, 7).GetValue<decimal>() == -9m, "Blank, zero and negative manual fields stay attached to their dates");
        Assert(sheet.Cell(oldRow, 12).Style.NumberFormat.Format == manualFormat, "Historical manual formatting is retained");
    }
}
