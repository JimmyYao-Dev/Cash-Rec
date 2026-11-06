using System.Globalization;
using System.Text.Json;
using CashRec;

try { Run(Path.GetFullPath(args[0]), Path.GetFullPath(args[1])); }
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run(string fixture, string evidencePath)
{
    var date = new DateTime(2026, 10, 25);
    var singleWorkspace = Inputs(fixture);
    using var singleFault = new DiagnosticFault(singleWorkspace);
    var single = ReportTaskRunner.Run(singleWorkspace, date, message => { if (message == "备份并提交报表") singleFault.Block(); });
    Assert(singleFault.Triggered && !single.Complete && !single.DiagnosticWritten && single.DiagnosticPath == null
        && single.DiagnosticAttemptedPath != null && single.Transaction?.Status == "Committed"
        && File.Exists(single.DailyOutput) && File.Exists(single.MonthlyOutput)
        && single.Diagnostics.Any(diagnostic => diagnostic.Code == DiagnosticCode.DiagnosticWriteFailed), "Committed single task retains its outputs and real transaction when the final record cannot be written");
    Assert(single.Transaction!.CommittedOutputs.Contains(single.DailyOutput!) && single.Transaction.CommittedOutputs.Contains(single.MonthlyOutput!), "Committed path inventory is preserved");
    singleFault.Restore();
    Assert(single.DiagnosticAttemptedPath != null && !File.Exists(single.DiagnosticAttemptedPath), "No failed task record is advertised as a written JSON file");

    var batchWorkspace = Inputs(fixture);
    using var batchFault = new DiagnosticFault(batchWorkspace);
    var batch = ReportBatchRunner.Run(batchWorkspace, new DateTime(2026, 10, 22), date,
        message => { if (!batchFault.Triggered && message.EndsWith("备份并提交报表", StringComparison.Ordinal)) batchFault.Block(); }, _ => batchFault.Restore());
    Assert(batch.Status == "Partial" && !batch.Complete && batch.DiagnosticWritten && File.Exists(batch.DiagnosticPath)
        && batch.Days[0].Status == "Partial" && !batch.Days[0].DiagnosticWritten && batch.Days[0].DiagnosticPath == null
        && batch.Days[0].Transaction?.Status == "Committed" && batch.Days[^1].Status == "Success", "Batch preserves publication and continues to a later valid date after task-record failure");
    Assert(batch.PublishedOutputs.Count == 3 && batch.PublishedOutputs.All(File.Exists), "All earlier and later published workbook paths survive");
    using (var json = JsonDocument.Parse(File.ReadAllText(batch.DiagnosticPath!)))
        Assert(json.RootElement.GetProperty("Days")[0].GetProperty("Transaction").GetProperty("Status").GetString() == "Committed"
            && json.RootElement.GetProperty("Days")[0].GetProperty("DiagnosticPath").ValueKind == JsonValueKind.Null, "Written batch index accurately records an unwritten per-day record");

    var summaryWorkspace = Inputs(fixture);
    using var summaryFault = new DiagnosticFault(summaryWorkspace);
    var summary = ReportBatchRunner.Run(summaryWorkspace, date, date, completed: _ => summaryFault.Block());
    Assert(summary.Status == "Partial" && !summary.Complete && !summary.DiagnosticWritten && summary.DiagnosticPath == null
        && summary.Days.Single().Status == "Success" && summary.Days.Single().Transaction?.Status == "Committed"
        && summary.PublishedOutputs.Count == 2 && summary.PublishedOutputs.All(File.Exists)
        && summary.Diagnostics.Any(diagnostic => diagnostic.Code == DiagnosticCode.DiagnosticWriteFailed), "A failed batch record cannot erase otherwise successful day state or claim overall success");
    summaryFault.Restore();
    Assert(File.Exists(summary.Days.Single().DiagnosticPath), "The successfully saved day record survives a summary-only failure");

    var developerWorkspace = Inputs(fixture);
    using var developerFault = new DiagnosticFault(developerWorkspace);
    var developer = ReportRunner.Run(developerWorkspace, date, message => { if (message == "准备基础报表") developerFault.Block(); });
    Assert(File.Exists(developer.OutputPath) && !developer.IsComplete && !developer.DiagnosticWritten && developer.DiagnosticPath == null,
        "The related developer-only daily entry also retains already committed output");
    developerFault.Restore();
    var evidence = new { AuxiliaryOnly = true, Single = single, Batch = batch, SummaryFailure = summary, DeveloperDaily = developer,
        SingleWorkspace = singleWorkspace, BatchWorkspace = batchWorkspace, SummaryWorkspace = summaryWorkspace, DeveloperWorkspace = developerWorkspace };
    Directory.CreateDirectory(Path.GetDirectoryName(evidencePath)!);
    File.WriteAllText(evidencePath, JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine("PASS: actual diagnostics-only I/O failures preserve committed report paths/transaction state, do not claim nonexistent records, continue later days, and report honest single/batch partial outcomes.");
    Console.WriteLine("Auxiliary evidence: " + evidencePath);
}
static string Inputs(string source)
{
    var target = Path.Combine(Path.GetTempPath(), "cashrec-diagnostic-" + Guid.NewGuid().ToString("N"));
    Copy(Path.Combine(source, "db"), Path.Combine(target, "db"));
    foreach (var directory in Directory.EnumerateDirectories(source).Where(path => DateTime.TryParseExact(Path.GetFileName(path), "yyyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _)))
        Copy(Path.Combine(directory, "Needed files"), Path.Combine(target, Path.GetFileName(directory), "Needed files"));
    return target;
}
static void Copy(string source, string target) { Directory.CreateDirectory(target); foreach (var file in Directory.EnumerateFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file))); foreach (var child in Directory.EnumerateDirectories(source)) Copy(child, Path.Combine(target, Path.GetFileName(child))); }
static void Assert(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }

sealed class DiagnosticFault(string workspace) : IDisposable
{
    readonly string path = Path.Combine(workspace, ".cashrec", "diagnostics");
    readonly string held = Path.Combine(workspace, ".cashrec", "diagnostics-held-" + Guid.NewGuid().ToString("N"));
    bool active;
    public bool Triggered { get; private set; }
    public void Block()
    {
        if (Triggered) return;
        Directory.Move(path, held); active = true;
        File.WriteAllText(path, "Synthetic diagnostics-only I/O fault"); Triggered = true;
    }
    public void Restore() { if (!active) return; File.Delete(path); Directory.Move(held, path); active = false; }
    public void Dispose() => Restore();
}
