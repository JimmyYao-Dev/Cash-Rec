using System.Security.Cryptography;
using CashRec;
using ClosedXML.Excel;

try { Run(args[0]); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run(string fixture)
{
    var workspace = Path.Combine(Path.GetTempPath(), "cashrec-rerun-" + Guid.NewGuid().ToString("N"));
    CopyDirectory(Path.GetFullPath(fixture), workspace);
    var current = new DateTime(2026, 10, 25); var previous = new DateTime(2026, 10, 22);
    Assert(ReportTaskRunner.Run(workspace, previous).Complete, "Previous complete report");
    var result = ReportTaskRunner.Run(workspace, current);
    Assert(result.Complete && result.Transaction?.Status == "Committed", "Initial paired report");
    var dailyPath = result.DailyOutput!; var monthPath = result.MonthlyOutput!;
    using (var book = new XLWorkbook(dailyPath))
    {
        var sheet = book.Worksheet("Equity"); var headers = Headers(sheet);
        foreach (var row in sheet.RowsUsed().Skip(1))
        {
            var branch = row.Cell(headers["Branch"]).GetString(); var kind = row.Cell(headers["Retail Type"]).GetString(); var currency = row.Cell(headers["Currency"]).GetString();
            var dep = row.Cell(headers["DEP"]);
            if (branch == "TMA" && kind == "Standard" && currency == "USD") { dep.Value = 30; dep.Style.NumberFormat.Format = "0.0000"; dep.Style.Font.Bold = true; dep.Style.Font.FontColor = XLColor.Blue; dep.Style.Fill.BackgroundColor = XLColor.Yellow; }
            if (branch == "TMA" && kind == "IB" && currency == "AUD") { dep.Value = -4; dep.Style.NumberFormat.Format = "0.000"; }
            if (branch == "TMG" && currency == "USC") { dep.Clear(XLClearOptions.Contents); dep.Style.NumberFormat.Format = "0.00000"; }
            if (branch == "TMG" && currency == "EUR") { dep.Value = 7; dep.Style.NumberFormat.Format = "0.000000"; }
        }
        sheet.Range(2, 1, sheet.LastRowUsed()!.RowNumber(), sheet.LastColumnUsed()!.ColumnNumber()).Sort(headers["Currency"], XLSortOrder.Descending, true);
        book.Worksheet("Bank").Position = 1; book.Save();
    }
    using (var book = new XLWorkbook(monthPath))
    {
        var sheet = book.Worksheet("Break down");
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > 1))
        {
            var date = row.Cell(1).GetDateTime().Date;
            row.Cell(7).Value = date == current ? -5 : 0;
            row.Cell(12).Value = date == current ? 0 : -3;
            row.Cell(13).Clear(XLClearOptions.Contents);
            row.Cell(14).Value = date == current ? -2.5 : 0;
            foreach (var column in new[] { 7, 12, 13, 14 }) row.Cell(column).Style.NumberFormat.Format = "0.0000";
        }
        foreach (var tab in book.Worksheets) tab.Range(MonthlyLayout.HeaderRows(tab.Name) + 1, 1, tab.LastRowUsed()!.RowNumber(), tab.LastColumnUsed()!.ColumnNumber()).Sort(1, XLSortOrder.Descending, false);
        book.Worksheet("Break down").Position = 1; book.Save();
    }
    var dailyBefore = Hash(dailyPath); var monthlyBefore = Hash(monthPath);
    // Real file lock, through the same GUI task coordinator used in production.
    using (var held = new FileStream(monthPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
    {
        var failed = ReportTaskRunner.Run(workspace, current);
        Assert(!failed.Complete && failed.DailyOutput == null && failed.MonthlyOutput == null && failed.Transaction?.Status == "NotCommitted", "Locked output cannot publish a mixed pair");
    }
    Assert(Hash(dailyPath) == dailyBefore && Hash(monthPath) == monthlyBefore, "Locked failure preserves exact workbook bytes");
    result = ReportTaskRunner.Run(workspace, current);
    Assert(result.Complete && result.Transaction?.Status == "Committed", string.Join(";", result.Diagnostics.Select(diagnostic => diagnostic.Message)));
    Assert(result.Transaction!.Backups.Select(Hash).Contains(dailyBefore) && result.Transaction.Backups.Select(Hash).Contains(monthlyBefore), "Both original byte snapshots are recoverable");
    var equity = BalanceWorkbookReader.ReadRows(dailyPath, "Equity");
    var standard = equity.Single(row => (string?)row["Branch"] == "TMA" && (string?)row["Retail Type"] == "Standard" && (string?)row["Currency"] == "USD");
    Assert((decimal)standard["DEP"]! == 30 && (decimal)standard["PNL"]! == -22, "DEP follows business key after row/tab reordering");
    Assert(equity.Single(row => (string?)row["Branch"] == "TMG" && (string?)row["Currency"] == "USC")["DEP"] == null, "True blank DEP survives");
    using (var book = new XLWorkbook(dailyPath))
    {
        var sheet = book.Worksheet("Equity"); var headers = Headers(sheet);
        var row = sheet.RowsUsed().Skip(1).Single(row => row.Cell(headers["Branch"]).GetString() == "TMA" && row.Cell(headers["Retail Type"]).GetString() == "Standard" && row.Cell(headers["Currency"]).GetString() == "USD");
        var dep = row.Cell(headers["DEP"]);
        Assert(dep.Style.NumberFormat.Format == "0.0000" && dep.Style.Font.Bold && dep.Style.Font.FontColor.Color.ToArgb() == XLColor.Blue.Color.ToArgb()
            && dep.Style.Fill.BackgroundColor.Color.ToArgb() == XLColor.Yellow.Color.ToArgb(), "Full manual DEP style survives lightweight style transfer");
    }
    using (var book = new XLWorkbook(monthPath))
    {
        Assert(book.Worksheets.Select(sheet => sheet.Name).SequenceEqual(MonthlyLayout.Sheets), "Named tabs are canonical without cache-sheet confusion");
        foreach (var tab in book.Worksheets)
        {
            var dates = tab.RowsUsed().Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(tab.Name)).Select(row => row.Cell(1).GetDateTime().Date).ToArray();
            Assert(dates.SequenceEqual(new[] { previous, current }), "Date rows sort once, with no duplicates");
        }
        var sheet = book.Worksheet("Break down");
        Assert(ExcelNumbers.Read(sheet.Cell("P2"), "prior Check") == -476 && ExcelNumbers.Read(sheet.Cell("P3"), "current Check") == -324.5m, "Every historical/current formula has the correct cache before Excel opens");
        Assert(sheet.Cell("G3").GetDouble() == -5 && sheet.Cell("L3").GetDouble() == 0 && sheet.Cell("M3").IsEmpty() && sheet.Cell("N3").GetDouble() == -2.5,
            "Dated manual blanks, zero and negatives survive reordering");
        Assert(sheet.Cell("M3").Style.NumberFormat.Format == "0.0000", "Even a blank manual cell retains its format");
    }
    CheckChangedKeys(workspace, current, previous);
    CheckLateRollback(workspace);
    Console.WriteLine("PASS: exact-byte locked failure, guarded late-failure compensation, recoverable backups, keyed DEP values/full styles, named tabs, chronological rows, preserved manual fields historical numeric caches, new-key explicit-zero basis and unmatched manual trace.");
    Console.WriteLine("Synthetic developer workspace: " + workspace);
}

static void CheckChangedKeys(string workspace, DateTime current, DateTime previous)
{
    var currentCsv = EquityPath(current);
    var records = CsvData.Read(currentCsv).ToList();
    records.RemoveAll(row => row["login"] == "SYN-E6");
    var added = new Dictionary<string, string>(records[0])
    {
        ["login"] = "NEW-SYN-IB-EUR", ["report_jurisdiction"] = "VFSC", ["is_ib"] = "true",
        ["currency"] = "EUR", ["positive_equity"] = "12.5",
    };
    records.Add(added); WriteCsv(currentCsv, records);
    var missing = ReportTaskRunner.Run(workspace, current);
    Assert(!missing.Complete && missing.Transaction?.Status == "Committed", "Unavailable prior new key is explicitly incomplete, not an assumed zero");
    var newRow = NewKey(missing.DailyOutput!);
    Assert(newRow["Movement"] == null && (decimal)newRow["DEP"]! == 0 && (decimal)BalanceWorkbookReader.ReadRows(missing.DailyOutput!, "Equity").First(row => row["Previous Date"] != null)["Previous Date"]! == (decimal)previous.ToOADate(),
        "New DEP defaults to zero but absent prior key stays blank on the same chosen prior date");
    Assert(missing.UnmatchedManualCount == 1 && missing.UnmatchedManualPath != null, "Disappeared saved manual key is separately traceable");
    var trace = CsvData.Read(missing.UnmatchedManualPath!).Single();
    Assert(trace["Key"] == "TMG/Standard/EUR" && trace["Value"] == "7" && trace["Field"] == "DEP" && trace["NumberFormat"] == "0.000000",
        "Unmatched CSV preserves original key, value and saved format");
    var previousRates = Path.Combine(workspace, previous.ToString("yyyy.MM.dd"), "Needed files", "Rate", "VFSC Rate.xlsx");
    File.Move(previousRates, previousRates + ".held");
    try
    {
        var withoutRaw = ReportRunner.Run(workspace, current);
        Assert(!withoutRaw.IsComplete && withoutRaw.OutputPath != null && NewKey(withoutRaw.OutputPath)["Movement"] == null,
            "Missing raw basis retains the same valid cache and missing key; it cannot jump to an older date");
        Assert(withoutRaw.Diagnostics.Any(diagnostic => diagnostic.Message.Contains("保留该日有效缓存")), "Cached fallback reason is visible");
    }
    finally { File.Move(previousRates + ".held", previousRates); }
    var priorRows = CsvData.Read(EquityPath(previous)).ToList();
    var explicitZero = new Dictionary<string, string>(added) { ["positive_equity"] = "0" };
    priorRows.Add(explicitZero); WriteCsv(EquityPath(previous), priorRows);
    var completed = ReportTaskRunner.Run(workspace, current);
    Assert(completed.Complete, string.Join(";", completed.Diagnostics.Select(diagnostic => diagnostic.Message)));
    newRow = NewKey(completed.DailyOutput!);
    Assert((decimal)newRow["Prev USD Equivalent"]! == 0 && (decimal)newRow["Movement"]! == 10 && (decimal)newRow["DEP"]! == 0,
        "Only an explicit same-date raw zero repairs a missing cached business key");
    Assert(completed.Diagnostics.Any(diagnostic => diagnostic.Message.Contains("缓存缺少当前业务键")), "Same-date cache repair provenance is explicit");
    return;
    string EquityPath(DateTime day) => Path.Combine(workspace, day.ToString("yyyy.MM.dd"), "Needed files", "MT4", "Risk", "Accounting Equity Report.csv");
    static Dictionary<string, object?> NewKey(string path) => BalanceWorkbookReader.ReadRows(path, "Equity").Single(row => (string?)row["Branch"] == "TMG" && (string?)row["Retail Type"] == "IB" && (string?)row["Currency"] == "EUR");
    static void WriteCsv(string path, List<Dictionary<string, string>> rows)
    {
        var headers = rows[0].Keys.ToArray();
        using var writer = new StreamWriter(path, false, new System.Text.UTF8Encoding(true));
        writer.WriteLine(string.Join(",", headers));
        foreach (var row in rows.AsEnumerable().Reverse()) writer.WriteLine(string.Join(",", headers.Select(header => "\"" + row[header].Replace("\"", "\"\"") + "\"")));
    }
}

static void CheckLateRollback(string workspace)
{
    var directory = Path.Combine(workspace, "late-failure"); Directory.CreateDirectory(directory);
    var a = Path.Combine(directory, "a.txt"); var b = Path.Combine(directory, "b.txt");
    File.WriteAllText(a, "original a"); File.WriteAllText(b, "original b");
    var originalA = Hash(a); var originalB = Hash(b);
    using var pa = Prepare(a, "replacement a"); using var pb = Prepare(b, "replacement b");
    var temporary = Path.Combine(directory, "third.tmp"); File.WriteAllText(temporary, "third");
    using var pc = new PreparedFile(Path.Combine(directory, "missing-parent", "c.txt"), temporary, Path.Combine(directory, "backups"));
    using var ga = OutputTransaction.Guard(a); using var gb = OutputTransaction.Guard(b);
    var status = new TransactionStatus();
    try { OutputTransaction.Commit([pa, pb, pc], status); throw new InvalidOperationException("Expected real late I/O failure"); }
    catch (IOException) { }
    Assert(status.Status == "RolledBack" && status.CommittedOutputs.Count == 0 && Hash(a) == originalA && Hash(b) == originalB, "Late third-output failure restores both first replacements byte-for-byte under actual guard lifecycle");
    Assert(status.Backups.Count == 2 && status.Backups.Select(Hash).Contains(originalA) && status.Backups.Select(Hash).Contains(originalB), "Rollback does not consume backup files");
    return;
    PreparedFile Prepare(string destination, string text) { var temp = OutputTransaction.TemporaryFor(destination); File.WriteAllText(temp, text); return new(destination, temp, Path.Combine(directory, "backups")); }
}
static Dictionary<string, int> Headers(IXLWorksheet sheet) => sheet.Row(1).CellsUsed().ToDictionary(cell => cell.GetString(), cell => cell.Address.ColumnNumber);
static void CopyDirectory(string source, string target) { Directory.CreateDirectory(target); foreach (var file in Directory.EnumerateFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file))); foreach (var child in Directory.EnumerateDirectories(source)) CopyDirectory(child, Path.Combine(target, Path.GetFileName(child))); }
static string Hash(string path) { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)); }
static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
