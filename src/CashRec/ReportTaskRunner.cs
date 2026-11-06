using System.Text.Json;

namespace CashRec;

/// <summary>The installed GUI's single task; daily and monthly states stay explicit.</summary>
public static class ReportTaskRunner
{
    public static ReportTaskResult Run(string workspace, DateTime date, Action<string>? progress = null)
    {
        var daily = ReportRunner.RunDetailed(workspace, date, progress);
        var monthly = MonthlyReportRunner.Run(daily);
        var diagnostics = daily.Result.Diagnostics.Concat(monthly.Diagnostics).ToList();
        var directory = Path.GetDirectoryName(daily.Result.DiagnosticPath)!;
        var path = Path.Combine(directory, date.ToString("yyyy.MM.dd") + "-task-" + DateTime.Now.ToString("HHmmssfff") + ".json");
        var result = new ReportTaskResult(daily.Result.OutputPath, monthly.OutputPath, path,
            monthly.Stage, monthly.StageComplete, daily.Result.IsComplete && monthly.StageComplete, diagnostics);
        File.WriteAllText(path, JsonSerializer.Serialize(new
        {
            Output = result.DailyOutput, result.DailyOutput, result.MonthlyOutput, result.MonthlyStage,
            result.MonthlyStageComplete, result.Complete, DailyComplete = daily.Result.IsComplete,
            Classification = ReportPolicy.Classification,
            result.Diagnostics,
        }, new JsonSerializerOptions { WriteIndented = true }));
        return result;
    }
}
