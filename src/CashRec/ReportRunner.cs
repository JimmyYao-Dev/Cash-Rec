using System.Reflection;
using System.Text.Json;

namespace CashRec;

public static class ReportRunner
{
    public static IEnumerable<IReportAdapter> Adapters() => Assembly.GetExecutingAssembly().GetTypes()
        .Where(type => !type.IsAbstract && typeof(IReportAdapter).IsAssignableFrom(type))
        .OrderBy(type => type.Name).Select(type => (IReportAdapter)Activator.CreateInstance(type)!);

    public static RunResult Run(string workspace, DateTime date, Action<string>? progress = null)
    {
        var check = WorkspaceInputs.Inspect(workspace, date);
        var diagnostics = check.Diagnostics;
        var diagnosticPath = DiagnosticPath(workspace, date);
        if (!check.CanRun) return Finish(null, diagnosticPath, false, diagnostics);
        try
        {
            progress?.Invoke("读取每日汇率");
            var context = new RunContext(date, check.DateDirectory, check.Inputs, check.MappingPath,
                DailyRates.Load(check.Inputs["Rates"], date), progress);
            var tables = new List<ReportTable>();
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
            if (diagnostics.Any(d => d.Severity == "Error")) return Finish(null, diagnosticPath, false, diagnostics);
            if (tables.Count == 0) return Finish(null, diagnosticPath, false, diagnostics);
            var output = Path.Combine(check.DateDirectory, "CashRec_Base.xlsx");
            var temporary = Path.Combine(check.DateDirectory, ".CashRec-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                progress?.Invoke("保存基础报表");
                WorkbookWriter.Write(temporary, tables);
                if (File.Exists(output))
                {
                    var backups = Path.Combine(check.DateDirectory, "backups");
                    Directory.CreateDirectory(backups);
                    File.Copy(output, Path.Combine(backups, "CashRec_Base-" + DateTime.Now.ToString("yyyyMMdd-HHmmssfff") + ".xlsx"));
                }
                File.Move(temporary, output, true);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
            return Finish(output, diagnosticPath, ready, diagnostics);
        }
        catch (Exception error)
        {
            diagnostics.Add(new("Error", "处理阶段", error.Message));
            return Finish(null, diagnosticPath, false, diagnostics);
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
        File.WriteAllText(path, JsonSerializer.Serialize(new { Output = output, Complete = complete, Diagnostics = diagnostics }, new JsonSerializerOptions { WriteIndented = true }));
        return new(output, path, complete, diagnostics);
    }
}
