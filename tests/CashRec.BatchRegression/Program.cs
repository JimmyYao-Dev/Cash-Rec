using System.Globalization;
using System.Security.Cryptography;
using CashRec;
using ClosedXML.Excel;

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        try { if (args[0] == "--layout") CheckLayout(); else CheckRange(Path.GetFullPath(args[0])); return 0; }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    static void CheckRange(string fixture)
    {
        var start = new DateTime(2026, 10, 31); var end = new DateTime(2026, 11, 3);
        var workspace = Inputs(fixture); Seed(workspace);
        string? octoberHash = null;
        var result = ReportBatchRunner.Run(workspace, start, end, completed: day => { if (day.Date == start) octoberHash = Hash(Month(workspace, 10)); });
        Assert(result.Complete && result.Status == "Success" && result.Counts["Success"] == 4 && result.Days.Select(day => day.Date).SequenceEqual(ReportBatchRunner.Dates(start, end)), "Inclusive ascending weekend/month boundary with distinct day tasks");
        Assert(result.PublishedOutputs.Count == 6 && result.Days.All(day => File.Exists(day.DiagnosticPath)), "All daily and monthly paths remain available");
        Assert(Hash(Month(workspace, 10)) == octoberHash, "New-month tasks do not change the prior month");
        var expected = new[] { -595m, -565m, -575m, -580m };
        for (var i = 0; i < result.Days.Count; i++)
        {
            var day = result.Days[i];
            Assert(Check(day.MonthlyOutput!, day.Date) == expected[i], "Independent dated Pending/Check arithmetic");
            var transactions = BalanceWorkbookReader.ReadRows(day.DailyOutput!, "TransactionExchange");
            Assert(transactions.Count == 28 && transactions.Sum(row => (decimal)row["USD Equivalent"]!) == 2661m
                && transactions.All(row => (DateTime)row["Date"]! == day.Date), "Prepared multi-date CSV transactions remain complete and belong to each selected D");
            var sourceDates = transactions.Select(row => row.GetValueOrDefault("Source Date")?.ToString() ?? "").ToArray();
            Assert(sourceDates.Any(value => value.Contains("2026-10-23")) && sourceDates.Any(value => value.Contains("2026-10-24")), "Both outside-folder CSV source dates are retained");
            Assert(BalanceWorkbookReader.ReadRows(day.DailyOutput!, "PSP").Sum(row => (decimal)row["USD Equivalent"]!) == 253m, "Each folder uses its own exact-date PSP column");
        }
        var repeated = ReportBatchRunner.Run(workspace, start, end);
        Assert(repeated.Complete, "Repeated range succeeds without duplicate dates");
        Dates(Month(workspace, 11), [start, start.AddDays(1), start.AddDays(2), end]);
        var monthlyHash = Hash(Month(workspace, 11)); var dailyHash = Hash(result.Days[^1].DailyOutput!);
        var taskCount = Directory.GetFiles(Path.Combine(workspace, ".cashrec", "diagnostics"), "*-task-*.json").Length;
        var reverse = ReportBatchRunner.Run(workspace, end, start);
        Assert(!reverse.Complete && reverse.Status == "Failed" && reverse.Days.Count == 0 && reverse.Diagnostics.Single().Code == DiagnosticCode.InvalidDateRange
            && Hash(Month(workspace, 11)) == monthlyHash && Hash(result.Days[^1].DailyOutput!) == dailyHash
            && Directory.GetFiles(Path.Combine(workspace, ".cashrec", "diagnostics"), "*-task-*.json").Length == taskCount, "Reverse range cannot run days or replace outputs");

        foreach (var scenario in new[] { "folder", "file", "psp-date", "psp-duplicate" })
        {
            var broken = Inputs(fixture); Seed(broken); var missingDate = start.AddDays(scenario == "folder" ? 1 : 2);
            if (scenario == "folder") Directory.Move(Path.Combine(broken, "2026.11.01"), Path.Combine(broken, "held-November-first"));
            else if (scenario == "file") File.Move(Path.Combine(broken, "2026.11.02", "Needed files", "MT4", "MT4", "Transaction Summary Report.csv"), Path.Combine(broken, "held-transactions.csv"));
            else
            {
                using var psp = new XLWorkbook(Path.Combine(broken, "2026.11.02", "Needed files", "PSP+BANK", "PSP Balance Summary.xlsx"));
                var sheet = psp.Worksheet("NOV.2026");
                if (scenario == "psp-date") sheet.Cell("G1").Clear(XLClearOptions.Contents);
                else sheet.Cell("E1").Value = missingDate;
                psp.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
            }
            var batch = ReportBatchRunner.Run(broken, start, end);
            var missing = batch.Days.Single(day => day.Date == missingDate);
            Assert(batch.Status == "Partial" && !batch.Complete && batch.Days.Count == 4 && batch.Days[^1].Status == "Success", "Other valid dates continue after " + scenario);
            Assert(missing.Status == (scenario == "psp-duplicate" ? "Failed" : "Missing") && missing.DailyOutput == null && missing.MonthlyOutput == null,
                "Missing date/file/PSP date differs from ambiguous invalid input");
            Assert(batch.PublishedOutputs.Contains(batch.Days[0].DailyOutput!), "Later failure retains earlier committed result paths");
            if (scenario == "psp-date") Assert(missing.Diagnostics.Any(diagnostic => diagnostic.Code == DiagnosticCode.MissingPspDate)
                && Check(batch.Days[^1].MonthlyOutput!, end) == -560m, "Missing PSP D is not borrowed from an adjacent column or invalid prior snapshot");
        }
        var onlyCurrent = Inputs(fixture, "2026.10.25");
        var noBasis = ReportBatchRunner.Run(onlyCurrent, new DateTime(2026, 10, 25), new DateTime(2026, 10, 26));
        Assert(noBasis.Status == "Partial" && noBasis.Days[0].Status == "Partial" && noBasis.Days[0].Blocked && noBasis.Days[1].Status == "Missing"
            && BalanceWorkbookReader.ReadRows(noBasis.Days[0].DailyOutput!, "Equity").All(row => row["Movement"] == null), "Unavailable same-policy prior stays explicitly blocked/partial with blank Movement");
        var direct = Inputs(fixture); Seed(direct);
        var fromDateFolder = ReportBatchRunner.Run(Path.Combine(direct, "2026.10.31"), start, end);
        Assert(fromDateFolder.Complete && fromDateFolder.Workspace == direct && fromDateFolder.Days.All(day => Path.GetFileName(Path.GetDirectoryName(day.DailyOutput!)) == day.Date.ToString("yyyy.MM.dd")),
            "Direct date-folder range resolves each sibling date exactly once");
        Assert(!WorkspaceInputs.Inspect(Path.Combine(direct, "2026.10.31"), end).CanRun, "Single-day direct-folder mismatch guard remains intact");
        Console.WriteLine("PASS: inclusive ordered range, complete CSV windows, exact PSP dates, one-time rollover, missing/failed continuation, blocked prior, honest aggregate state, repeated upsert, reverse no-write and dated-folder resolution.");
        Console.WriteLine("Synthetic developer workspace: " + workspace);
    }

    static void CheckLayout()
    {
        using var form = new MainForm(); form.Show(); Application.DoEvents();
        T Find<T>(string name) where T : Control => (T)form.Controls.Find(name, true).Single();
        var start = Find<DateTimePicker>("businessDatePicker"); var end = Find<DateTimePicker>("endDatePicker"); var range = Find<CheckBox>("rangeCheckBox");
        Assert(!range.Checked && !end.Visible, "Single day remains the initial operation");
        start.Value = new DateTime(2026, 10, 31); range.Checked = true; Application.DoEvents();
        Assert(end.Visible && end.Value.Date == start.Value.Date && !Find<Button>("importButton").Enabled, "Enabling a range cannot silently choose a system-today end date or ambiguous multi-day import");
        foreach (var size in new[] { form.Size, form.MinimumSize })
        {
            form.Size = size; Application.DoEvents();
            foreach (var name in new[] { "businessDatePicker", "rangeCheckBox", "endDatePicker", "checkButton", "runButton", "openButton" })
            {
                var control = Find<Control>(name);
                Assert(control.Parent!.ClientRectangle.Contains(control.Bounds), "Visible action remains inside its parent at " + size + ": " + name);
            }
            Assert(Find<TextBox>("statusText").Height > 60, "Status remains usable at minimum size");
        }
        range.Checked = false; Assert(Find<Button>("importButton").Enabled && !end.Visible, "Returning to single day restores its import action");
        form.Close(); Console.WriteLine("PASS: developer-only default/minimum layout, range toggle and dates; installed GUI business acceptance is separate.");
    }

    static string Inputs(string source, string? only = null)
    {
        var target = Path.Combine(Path.GetTempPath(), "cashrec-batch-" + Guid.NewGuid().ToString("N"));
        CopyDirectory(Path.Combine(source, "db"), Path.Combine(target, "db"));
        foreach (var directory in Directory.EnumerateDirectories(source).Where(path => DateTime.TryParseExact(Path.GetFileName(path), "yyyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _) && (only == null || Path.GetFileName(path) == only)))
            CopyDirectory(Path.Combine(directory, "Needed files"), Path.Combine(target, Path.GetFileName(directory), "Needed files"));
        return target;
    }
    static void Seed(string workspace) { var result = ReportTaskRunner.Run(workspace, new DateTime(2026, 10, 25)); Assert(result.Complete, "Prepared prior month seed: " + string.Join(";", result.Diagnostics.Select(diagnostic => diagnostic.Message))); }
    static void Dates(string path, DateTime[] expected)
    {
        using var book = new XLWorkbook(path);
        foreach (var sheet in book.Worksheets)
            Assert(sheet.RowsUsed(XLCellsUsedOptions.Contents).Where(row => row.RowNumber() > (sheet.Name == "Break down" ? 1 : 2)).Select(row => row.Cell(1).GetDateTime().Date).SequenceEqual(expected), "Repeated batch keeps unique dated row set");
    }
    static decimal Check(string path, DateTime date)
    {
        using var book = new XLWorkbook(path); var sheet = book.Worksheet("Break down");
        var row = sheet.RowsUsed().Single(row => row.RowNumber() > 1 && row.Cell(1).TryGetValue<DateTime>(out var actual) && actual.Date == date);
        return ExcelNumbers.Read(row.Cell(16), "Check before Excel opens");
    }
    static string Month(string workspace, int month) => Path.Combine(workspace, "monthly", $"CashRec_Breakdown_2026.{month:00}.xlsx");
    static void CopyDirectory(string source, string target) { Directory.CreateDirectory(target); foreach (var file in Directory.EnumerateFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file))); foreach (var child in Directory.EnumerateDirectories(source)) CopyDirectory(child, Path.Combine(target, Path.GetFileName(child))); }
    static string Hash(string path) { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)); }
    static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
}
