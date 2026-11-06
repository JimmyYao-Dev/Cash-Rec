using System.Text.Json;

namespace CashRec;

internal sealed class PreparedMonthlyOutput(string destination, string temporaryPath, MonthlyReport report) : IDisposable
{
    public PreparedFile File { get; } = new(destination, temporaryPath, Path.Combine(Path.GetDirectoryName(destination)!, "backups"));
    public string Destination => File.Destination;
    public string TemporaryPath => File.TemporaryPath;
    public MonthlyReport Report { get; } = report;
    public void Dispose() => File.Dispose();
}

internal static class MonthlyReportRunner
{
    internal static PreparedMonthlyOutput Prepare(RunContext context, IList<ReportTable> tables)
    {
        context.Progress?.Invoke("汇总月度余额与在途结算");
        var report = MonthlyBalanceCalculator.Calculate(context, tables);
        context.Progress?.Invoke("汇总客户净入出金与批发净额");
        MonthlyFlowCalculator.Apply(report, tables, context.Progress);
        context.Progress?.Invoke("计算 Break down 资金差额");
        MonthlyBreakdownCalculator.Apply(report);
        var directory = Path.Combine(Directory.GetParent(context.DateDirectory)!.FullName, "monthly");
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, $"CashRec_Breakdown_{context.BusinessDate:yyyy.MM}.xlsx");
        var existing = File.Exists(destination) ? destination : null;
        context.Progress?.Invoke("计算结算佣金与 Check");
        MonthlySettlementCalculator.Apply(report, tables, MonthlyWorkbookWriter.ReadManualInputs(existing, context.BusinessDate), context.Progress);
        var temporary = Path.Combine(directory, ".CashRec-Monthly-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            MonthlyWorkbookWriter.Write(temporary, existing, report);
            return new(destination, temporary, report);
        }
        catch { if (File.Exists(temporary)) File.Delete(temporary); throw; }
    }

    internal static PreparedFile PrepareProvenance(PreparedMonthlyOutput output)
    {
        var directory = Path.Combine(Path.GetDirectoryName(output.Destination)!, ".cashrec");
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, output.Report.Date.ToString("yyyy.MM.dd") + ".json");
        var temporary = OutputTransaction.TemporaryFor(destination);
        File.WriteAllText(temporary, JsonSerializer.Serialize(new
        {
            Date = output.Report.Date, PreviousDate = output.Report.PreviousDate,
            MonthlyStage = MonthlyStage.Current, MonthlyStageComplete = output.Report.StageComplete,
            Classification = ReportPolicy.Classification, output.Report.Provenance, output.Report.Diagnostics,
        }, new JsonSerializerOptions { WriteIndented = true }));
        return new(destination, temporary, Path.Combine(directory, "backups"));
    }
}
