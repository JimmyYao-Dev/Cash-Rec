using System.Reflection;

namespace CashRec;

public static class ReportRunner
{
    public static IEnumerable<IReportAdapter> Adapters() => Assembly.GetExecutingAssembly().GetTypes()
        .Where(type => !type.IsAbstract && typeof(IReportAdapter).IsAssignableFrom(type))
        .OrderBy(type => type.Name).Select(type => (IReportAdapter)Activator.CreateInstance(type)!);

    public static RunResult Run(string workspace, DateTime date, Action<string>? progress = null)
    {
        var daily = Prepare(workspace, date, progress);
        using var output = daily.PreparedOutput;
        using var unmatched = daily.Context == null ? null : UnmatchedManualReport.Prepare(daily.Context.DateDirectory,
            Guid.NewGuid().ToString("N"), daily.Context.UnmatchedManual);
        if (output == null) return Finish(null, daily.Result.DiagnosticPath!, false, daily.Result.Diagnostics);
        try
        {
            OutputTransaction.Commit(unmatched == null ? [output] : [output, unmatched], new TransactionStatus());
            return Finish(output.Destination, daily.Result.DiagnosticPath!, daily.Result.IsComplete, daily.Result.Diagnostics);
        }
        catch (Exception error)
        {
            daily.Result.Diagnostics.Add(new("Error", "保存基础报表", error.Message));
            return Finish(null, daily.Result.DiagnosticPath!, false, daily.Result.Diagnostics);
        }
    }

    internal static DailyRunArtifacts Prepare(string workspace, DateTime date, Action<string>? progress = null, PreflightResult? preflight = null)
    {
        var check = preflight ?? WorkspaceInputs.Inspect(workspace, date);
        var diagnostics = check.Diagnostics;
        var diagnosticPath = DiagnosticPath(workspace, date);
        RunContext? context = null;
        var tables = new List<ReportTable>();
        PreparedFile? prepared = null;
        DailyRunArtifacts Complete(bool complete) =>
            new(new(null, diagnosticPath, complete, diagnostics), context, tables, prepared);
        if (!check.CanRun) return Complete(false);
        try
        {
            progress?.Invoke("读取每日汇率");
            context = new RunContext(date, check.DateDirectory, check.Inputs, check.MappingPath,
                DailyRates.Load(check.Inputs["Rates"], date), progress);
            var adapters = Adapters().ToList();
            foreach (var adapter in adapters)
            {
                progress?.Invoke("处理 " + adapter.Name);
                try
                {
                    var table = adapter.Process(context);
                    tables.Add(table);
                    diagnostics.AddRange(table.Diagnostics);
                }
                catch (Exception error)
                { diagnostics.Add(new("Error", adapter.Name, error.Message)); }
            }
            var finalizers = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => !type.IsAbstract && typeof(IReportFinalizer).IsAssignableFrom(type))
                .Select(type => (IReportFinalizer)Activator.CreateInstance(type)!).ToList();
            foreach (var finalizer in finalizers) finalizer.FinalizeReport(context, tables, diagnostics);
            var ready = WorkbookWriter.SheetNames.All(name => tables.Any(t => t.Name == name)) && finalizers.Count > 0
                && !diagnostics.Any(d => d.Severity == "Incomplete");
            if (!ready) diagnostics.Add(new("Warning", "基础报表", "本次基础报表尚不完整；请查看已生成的表与诊断记录"));
            if (diagnostics.Any(d => d.Severity == "Error")) return Complete(false);
            if (tables.Count == 0) return Complete(false);
            var output = Path.Combine(check.DateDirectory, "CashRec_Base.xlsx");
            prepared = new(output, OutputTransaction.TemporaryFor(output), Path.Combine(check.DateDirectory, "backups"));
            progress?.Invoke("准备基础报表");
            WorkbookWriter.Write(prepared.TemporaryPath, tables);
            return Complete(ready);
        }
        catch (Exception error)
        {
            diagnostics.Add(new("Error", "处理阶段", error.Message));
            prepared?.Dispose(); prepared = null;
            return Complete(false);
        }
    }

    public static string DiagnosticPath(string workspace, DateTime date)
    {
        var directory = Path.Combine(Path.GetFullPath(workspace), ".cashrec", "diagnostics");
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, date.ToString("yyyy.MM.dd") + "-" + DateTime.Now.ToString("HHmmssfff") + ".json");
    }
    static RunResult Finish(string? output, string path, bool complete, List<Diagnostic> diagnostics)
    {
        try
        {
            DiagnosticRecords.Write(path, new { Output = output, Complete = complete, DiagnosticPath = path, DiagnosticWritten = true, DiagnosticAttemptedPath = path, Diagnostics = diagnostics });
            return new(output, path, complete, diagnostics, DiagnosticAttemptedPath: path);
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            diagnostics.Add(new("Error", "任务记录", $"任务记录未保存：{path}；{error.Message}", DiagnosticCode.DiagnosticWriteFailed));
            return new(output, null, false, diagnostics, false, path);
        }
    }
}
