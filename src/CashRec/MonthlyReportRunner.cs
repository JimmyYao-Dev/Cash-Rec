using System.Text.Json;

namespace CashRec;

internal sealed class PreparedMonthlyOutput(string destination, string temporaryPath, MonthlyReport report) : IDisposable
{
    public string Destination { get; } = destination;
    public string TemporaryPath { get; } = temporaryPath;
    public MonthlyReport Report { get; } = report;
    public void Dispose() { if (File.Exists(TemporaryPath)) File.Delete(TemporaryPath); }
}

internal static class MonthlyReportRunner
{
    public static MonthlyRunResult Run(DailyRunArtifacts daily)
    {
        if (daily.Result.OutputPath == null || daily.Context == null)
            return new(null, MonthlyStage.Current, false, [new("Incomplete", "月度报表", "基础报表未生成，本次未更新月度结果")]);
        try
        {
            using var prepared = Prepare(daily.Context, daily.Tables);
            Commit(prepared);
            var diagnostics = prepared.Report.Diagnostics;
            diagnostics.Add(new("Info", "月度报表", prepared.Report.StageComplete
                ? "已生成 TMA、TMG、Total 的余额、DIFF、SETTLE PNL 和客户净额"
                : "已生成当前余额和客户净额；部分上期数据不可用，对应 DIFF 留空"));
            return new(prepared.Destination, MonthlyStage.Current, prepared.Report.StageComplete, diagnostics);
        }
        catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException or InvalidOperationException)
        {
            return new(null, MonthlyStage.Current, false, [new("Error", "月度报表", error.Message)]);
        }
    }

    internal static PreparedMonthlyOutput Prepare(RunContext context, IList<ReportTable> tables)
    {
        context.Progress?.Invoke("汇总月度余额与在途结算");
        var report = MonthlyBalanceCalculator.Calculate(context, tables);
        context.Progress?.Invoke("汇总客户净入出金与批发净额");
        MonthlyFlowCalculator.Apply(report, tables, context.Progress);
        var directory = Path.Combine(Directory.GetParent(context.DateDirectory)!.FullName, "monthly");
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, $"CashRec_Breakdown_{context.BusinessDate:yyyy.MM}.xlsx");
        var temporary = Path.Combine(directory, ".CashRec-Monthly-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            MonthlyWorkbookWriter.Write(temporary, File.Exists(destination) ? destination : null, report);
            return new(destination, temporary, report);
        }
        catch { if (File.Exists(temporary)) File.Delete(temporary); throw; }
    }

    internal static void Commit(PreparedMonthlyOutput output)
    {
        var directory = Path.GetDirectoryName(output.Destination)!;
        var provenanceDirectory = Path.Combine(directory, ".cashrec");
        Directory.CreateDirectory(provenanceDirectory);
        // Preparation is separate from replacement so a later task-level transaction can
        // coordinate daily/monthly outputs without changing the calculation contract.
        if (File.Exists(output.Destination))
        {
            var backups = Path.Combine(directory, "backups");
            Directory.CreateDirectory(backups);
            File.Copy(output.Destination, Path.Combine(backups, Path.GetFileNameWithoutExtension(output.Destination)
                + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx"));
        }
        File.Move(output.TemporaryPath, output.Destination, true);
        File.WriteAllText(Path.Combine(provenanceDirectory, output.Report.Date.ToString("yyyy.MM.dd") + ".json"),
            JsonSerializer.Serialize(new
            {
                Date = output.Report.Date, PreviousDate = output.Report.PreviousDate,
                MonthlyStage = MonthlyStage.Current, MonthlyStageComplete = output.Report.StageComplete,
                Classification = ReportPolicy.Classification,
                output.Report.Provenance, output.Report.Diagnostics,
            }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
