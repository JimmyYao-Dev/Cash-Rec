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
        diagnostics.Add(new("Pending", "月度阶段", "Break down 差额、结算佣金及 Check 尚未生成；本次结果不是完整月度核对"));
        var directory = Path.GetDirectoryName(daily.Result.DiagnosticPath)!;
        var path = Path.Combine(directory, date.ToString("yyyy.MM.dd") + "-task-" + DateTime.Now.ToString("HHmmssfff") + ".json");
        var result = new ReportTaskResult(daily.Result.OutputPath, monthly.OutputPath, path,
            monthly.Stage, monthly.StageComplete, false, diagnostics);
        File.WriteAllText(path, JsonSerializer.Serialize(new
        {
            Output = result.DailyOutput, result.DailyOutput, result.MonthlyOutput, result.MonthlyStage,
            result.MonthlyStageComplete, result.Complete, DailyComplete = daily.Result.IsComplete,
            result.Diagnostics,
        }, new JsonSerializerOptions { WriteIndented = true }));
        return result;
    }
}
