using System.Security.Cryptography;
using CashRec;
using ClosedXML.Excel;

try { Run(Path.GetFullPath(args[0])); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run(string fixture)
{
    var workspace = Clone(fixture);
    var october = new DateTime(2026, 10, 31); var first = new DateTime(2026, 11, 2); var second = first.AddDays(1);
    Assert(ReportTaskRunner.Run(workspace, new DateTime(2026, 10, 25)).Complete, "Mid-month bootstrap works without a previous monthly file");
    var result = ReportTaskRunner.Run(workspace, october);
    Assert(result.Complete, "October 31 append");
    var source = result.MonthlyOutput!;
    using (var book = new XLWorkbook(source))
    {
        var sheet = book.Worksheet("Break down"); var row = DateRow(sheet, october);
        sheet.Cell(row, 7).Value = -5; sheet.Cell(row, 12).Value = 0;
        sheet.Cell(row, 13).Clear(XLClearOptions.Contents); sheet.Cell(row, 14).Value = -2.5;
        foreach (var column in new[] { 7, 12, 13, 14 }) { sheet.Cell(row, column).Style.NumberFormat.Format = "0.0000"; sheet.Cell(row, column).Style.Font.Bold = true; }
        foreach (var tab in book.Worksheets)
        {
            tab.Cell(80, 1).Value = "Grand Total"; tab.Cell(80, 2).Value = 999999;
            tab.Cell(81, 1).Value = "Operator note"; tab.Cell(100, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
        }
        book.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
    }
    var sourceHash = Hash(source);
    result = ReportTaskRunner.Run(workspace, first);
    Assert(result.Complete, Diagnostics(result));
    CheckDates(result.MonthlyOutput!, [october, first]);
    using (var book = new XLWorkbook(result.MonthlyOutput!))
    {
        var sheet = book.Worksheet("Break down"); var carry = DateRow(sheet, october); var current = DateRow(sheet, first);
        Assert(ExcelNumbers.Read(sheet.Cell(carry, 16), "carried Check") == -587.5m && ExcelNumbers.Read(sheet.Cell(current, 16), "current Check") == -545m,
            "Carried and new-row Check caches are correct before Excel opens");
        Assert(sheet.Cell(carry, 7).GetDouble() == -5 && sheet.Cell(carry, 13).IsEmpty() && sheet.Cell(carry, 13).Style.NumberFormat.Format == "0.0000" && sheet.Cell(carry, 13).Style.Font.Bold,
            "Baseline blank/value/full format survives cleanup");
        Assert(book.CustomProperties.Single(property => property.Name == "CashRecMonthlyBaselineDate").Value.ToString() == "2026.10.31", "Explicit initialization baseline");
        Assert(!book.Worksheets.SelectMany(tab => tab.CellsUsed(XLCellsUsedOptions.Contents)).Any(cell => cell.GetString() is "Grand Total" or "Operator note"), "Trailing totals/notes are not carried");
        Assert(book.Worksheets.All(tab => tab.LastRowUsed(XLCellsUsedOptions.All)!.RowNumber() < 80), "Trailing style-only rows are removed from the new-month copy");
    }
    result = ReportTaskRunner.Run(workspace, second);
    Assert(result.Complete, Diagnostics(result));
    CheckDates(result.MonthlyOutput!, [october, first, second]);
    using (var book = new XLWorkbook(result.MonthlyOutput!))
        Assert(ExcelNumbers.Read(book.Worksheet("Break down").Cell(DateRow(book.Worksheet("Break down"), second), 16), "next Check") == -580m,
            "Following day appends without reinitializing the month");
    Assert(Hash(source) == sourceHash, "Prior month remains byte-for-byte unchanged");
    foreach (var scenario in new[] { "missing", "inconsistent", "empty", "old-policy" })
    {
        var bad = Clone(workspace); var old = MonthPath(bad, 10); File.Delete(MonthPath(bad, 11));
        if (scenario == "missing") File.Move(old, old + ".held");
        else
        {
            using var book = new XLWorkbook(old);
            if (scenario == "old-policy") book.CustomProperties.Single(property => property.Name == ReportPolicy.ClassificationDateProperty(october)).Value = "legacy";
            else foreach (var tab in book.Worksheets)
                foreach (var row in tab.RowsUsed().Where(row => row.RowNumber() > (tab.Name == "Break down" ? 1 : 2)).ToList())
                    if (scenario == "empty" || tab.Name == "Break down" && row.Cell(1).TryGetValue<DateTime>(out var date) && date.Date == october)
                        row.Clear(XLClearOptions.Contents);
            book.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }
        var existingDaily = Path.Combine(bad, "2026.11.02", "CashRec_Base.xlsx"); var before = Hash(existingDaily);
        var failed = ReportTaskRunner.Run(bad, first);
        Assert(!failed.Complete && failed.DailyOutput == null && failed.MonthlyOutput == null && Hash(existingDaily) == before && !File.Exists(MonthPath(bad, 11)), scenario + " refuses invalid baseline without damaging results");
        if (scenario == "old-policy") Assert(Diagnostics(failed).Contains("新Comment规则"), "Stale classification asks for an explicit baseline rerun");
    }
    CheckLateBasis(fixture, workspace, true);
    CheckLateBasis(fixture, workspace, false);
    Console.WriteLine("PASS: mid-month append, one-time four-sheet rollover, trailing-row exclusion, source-byte preservation, carry manual styles/caches, invalid baselines and partial-output baseline repair.");
    Console.WriteLine("Synthetic developer workspace: " + workspace);
}

static void CheckLateBasis(string fixture, string prepared, bool monthFirst)
{
    var workspace = Path.Combine(Path.GetTempPath(), "cashrec-late-basis-" + Guid.NewGuid().ToString("N"));
    CopyDirectory(Path.Combine(fixture, "db"), Path.Combine(workspace, "db"));
    var date = new DateTime(2026, 11, monthFirst ? 1 : 2);
    CopyDirectory(Path.Combine(fixture, "2026.11.02", "Needed files"), Path.Combine(workspace, date.ToString("yyyy.MM.dd"), "Needed files"));
    if (monthFirst) PrepareFirstDate(workspace, date);
    var partial = ReportTaskRunner.Run(workspace, date);
    Assert(!partial.Complete && partial.MonthlyOutput != null, "No-prior first use remains an explicit partial output");
    using (var book = new XLWorkbook(partial.MonthlyOutput!))
    {
        var sheet = book.Worksheet("Break down"); var row = DateRow(sheet, date); sheet.Cell(row, 7).Value = -4;
        sheet.Cell(row, 7).Style.NumberFormat.Format = "0.00000"; book.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
    }
    var next = date.AddDays(1);
    CopyDirectory(Path.Combine(fixture, "2026.11.03", "Needed files"), Path.Combine(workspace, next.ToString("yyyy.MM.dd"), "Needed files"));
    if (monthFirst) PrepareFirstDate(workspace, next);
    var follow = ReportTaskRunner.Run(workspace, next);
    Assert(monthFirst ? follow.MonthlyOutput == null && !follow.Complete : follow.Complete,
        "Month-first pending carry cannot be hidden by a later same-month basis; mid-month bootstrap remains available");
    foreach (var directory in Directory.EnumerateDirectories(fixture).Where(path => Path.GetFileName(path).StartsWith("2026.10.", StringComparison.Ordinal)))
        CopyDirectory(directory, Path.Combine(workspace, Path.GetFileName(directory)));
    Directory.CreateDirectory(Path.Combine(workspace, "monthly")); File.Copy(MonthPath(prepared, 10), MonthPath(workspace, 10));
    var repaired = ReportTaskRunner.Run(workspace, date);
    Assert(repaired.Complete, Diagnostics(repaired));
    CheckDates(repaired.MonthlyOutput!, monthFirst ? [new DateTime(2026, 10, 31), date] : [new DateTime(2026, 10, 31), date, next]);
    using var result = new XLWorkbook(repaired.MonthlyOutput!);
    Assert(result.CustomProperties.Single(property => property.Name == ReportPolicy.ClassificationDateProperty(new DateTime(2026, 10, 31))).Value.ToString() == ReportPolicy.Classification
        && result.CustomProperties.Single(property => property.Name == "CashRecClassification").Value.ToString() == ReportPolicy.Classification,
        "A verified late-carried source date retains its policy marker and cannot produce a spurious Mixed warning");
    Assert(!repaired.Diagnostics.Any(diagnostic => diagnostic.Source == "历史交易分类"), "No current-policy carry is mislabeled as unrecomputed history");
    var manual = result.Worksheet("Break down").Cell(DateRow(result.Worksheet("Break down"), date), 7);
    Assert(manual.GetDouble() == -4 && manual.Style.NumberFormat.Format == "0.00000", "Adding a late baseline preserves existing new-month manual input and format");
}

static void PrepareFirstDate(string workspace, DateTime date)
{
    var needed = Path.Combine(workspace, date.ToString("yyyy.MM.dd"), "Needed files");
    using (var rates = new XLWorkbook(Path.Combine(needed, "Rate", "VFSC Rate.xlsx")))
    { foreach (var row in rates.Worksheets.First().RowsUsed().Skip(1)) row.Cell(1).Value = date; rates.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true }); }
    using (var psp = new XLWorkbook(Path.Combine(needed, "PSP+BANK", "PSP Balance Summary.xlsx")))
    { psp.Worksheet("NOV.2026").Cell("G1").Value = date; psp.Worksheet("NOV.2026").Cell("H1").Value = date; psp.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true }); }
    using var bank = new XLWorkbook(Path.Combine(needed, "PSP+BANK", "BANK BD.xlsx"));
    var snapshot = bank.CustomProperties.SingleOrDefault(property => property.Name == "CashRecSnapshotDate");
    if (snapshot != null) snapshot.Value = date.ToString("yyyy.MM.dd");
    bank.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
}
static string MonthPath(string workspace, int month) => Path.Combine(workspace, "monthly", $"CashRec_Breakdown_2026.{month:00}.xlsx");
static void CheckDates(string path, DateTime[] expected)
{
    using var book = new XLWorkbook(path);
    foreach (var sheet in book.Worksheets)
    {
        var dates = sheet.RowsUsed(XLCellsUsedOptions.Contents).Where(row => row.RowNumber() > (sheet.Name == "Break down" ? 1 : 2))
            .Where(row => row.Cell(1).Value.Type != XLDataType.Blank).Select(row => row.Cell(1).GetDateTime().Date).ToArray();
        Assert(dates.SequenceEqual(expected), sheet.Name + " has only the expected unique chronological business dates");
    }
}
static int DateRow(IXLWorksheet sheet, DateTime date) => sheet.RowsUsed().Single(row => row.RowNumber() > (sheet.Name == "Break down" ? 1 : 2) && row.Cell(1).TryGetValue<DateTime>(out var actual) && actual.Date == date).RowNumber();
static string Clone(string source) { var destination = Path.Combine(Path.GetTempPath(), "cashrec-rollover-" + Guid.NewGuid().ToString("N")); CopyDirectory(source, destination); return destination; }
static void CopyDirectory(string source, string target) { Directory.CreateDirectory(target); foreach (var file in Directory.EnumerateFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file))); foreach (var child in Directory.EnumerateDirectories(source)) CopyDirectory(child, Path.Combine(target, Path.GetFileName(child))); }
static string Hash(string path) { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)); }
static string Diagnostics(ReportTaskResult result) => string.Join(";", result.Diagnostics.Select(diagnostic => diagnostic.Message));
static void Assert(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
