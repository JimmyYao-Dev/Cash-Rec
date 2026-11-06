namespace CashRec;

/// <summary>The installed GUI's single task; daily and monthly states stay explicit.</summary>
public static class ReportTaskRunner
{
    public static ReportTaskResult Run(string workspace, DateTime date, Action<string>? progress = null)
    {
        var check = WorkspaceInputs.Inspect(workspace, date);
        var diagnostics = new List<Diagnostic>();
        var transaction = new TransactionStatus();
        var directory = Path.GetDirectoryName(ReportRunner.DiagnosticPath(workspace, date))!;
        var path = Path.Combine(directory, date.ToString("yyyy.MM.dd") + "-task-" + transaction.Id + ".json");
        DailyRunArtifacts? daily = null;
        PreparedMonthlyOutput? monthly = null;
        PreparedFile? trace = null, unmatched = null;
        var originals = new List<FileStream>();
        var retained = new List<string>();
        FileStream? runLock = null;
        var entries = new List<UnmatchedManualEntry>();
        try
        {
            if (check.CanRun)
            {
                var monthlyDirectory = Path.Combine(Directory.GetParent(check.DateDirectory)!.FullName, "monthly");
                Directory.CreateDirectory(monthlyDirectory);
                runLock = new FileStream(Path.Combine(monthlyDirectory, $".CashRec-{date:yyyy.MM}.run.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                foreach (var target in new[] { Path.Combine(check.DateDirectory, "CashRec_Base.xlsx"), Path.Combine(monthlyDirectory, $"CashRec_Breakdown_{date:yyyy.MM}.xlsx") })
                {
                    if (!File.Exists(target)) continue;
                    retained.Add(target); originals.Add(OutputTransaction.Guard(target));
                }
            }
            daily = ReportRunner.Prepare(workspace, date, progress, check);
            diagnostics.AddRange(daily.Result.Diagnostics);
            if (daily.PreparedOutput != null && daily.Context != null)
            {
                entries.AddRange(daily.Context.UnmatchedManual);
                monthly = MonthlyReportRunner.Prepare(daily.Context, daily.Tables);
                diagnostics.AddRange(monthly.Report.Diagnostics);
                entries.AddRange(monthly.Report.UnmatchedManual);
                trace = MonthlyReportRunner.PrepareProvenance(monthly);
                unmatched = UnmatchedManualReport.Prepare(check.DateDirectory, transaction.Id, entries);
                var outputs = new List<PreparedFile> { monthly.File, daily.PreparedOutput, trace };
                if (unmatched != null) outputs.Add(unmatched);
                progress?.Invoke("备份并提交报表");
                OutputTransaction.Commit(outputs, transaction);
                diagnostics.Add(new("Info", "报表输出", "基础表和月度结果已一起提交；生成完成不等于人工核对通过"));
                if (unmatched != null) diagnostics.Add(new("Warning", "未匹配人工值", $"{entries.Count}项人工键未匹配，已保留清单：{unmatched.Destination}"));
            }
        }
        catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException or InvalidOperationException)
        {
            diagnostics.AddRange(daily == null ? check.Diagnostics : []);
            diagnostics.Add(new("Error", "报表任务", error.Message));
            if (retained.Count > 0 && transaction.Status != "RollbackFailed")
                diagnostics.Add(new("Info", "原结果", "未发布新报表；原结果已保留：" + string.Join("；", retained)));
        }
        finally
        {
            daily?.PreparedOutput?.Dispose(); monthly?.Dispose(); trace?.Dispose(); unmatched?.Dispose();
            foreach (var original in originals) original.Dispose();
            runLock?.Dispose();
        }
        var committed = transaction.Status == "Committed";
        var result = new ReportTaskResult(committed ? daily!.PreparedOutput!.Destination : null, committed ? monthly!.Destination : null, path,
            MonthlyStage.Current, committed && monthly!.Report.StageComplete,
            committed && daily!.Result.IsComplete && monthly!.Report.StageComplete, diagnostics, transaction,
            committed ? unmatched?.Destination : null, entries.Count, DiagnosticAttemptedPath: path);
        try
        {
            DiagnosticRecords.Write(path, new
            {
                Output = result.DailyOutput, result.DailyOutput, result.MonthlyOutput, result.MonthlyStage,
                result.MonthlyStageComplete, result.Complete, DailyComplete = committed && daily!.Result.IsComplete,
                Classification = ReportPolicy.Classification,
                result.Transaction, result.UnmatchedManualPath, result.UnmatchedManualCount,
                result.DiagnosticPath, result.DiagnosticWritten, result.DiagnosticAttemptedPath,
                RetainedOutputs = !committed ? retained : [], result.Diagnostics,
            });
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            diagnostics.Add(new("Error", "任务记录", $"{(committed ? "报表已提交；" : "")}任务记录未保存：{path}；{error.Message}", DiagnosticCode.DiagnosticWriteFailed));
            result = result with { Complete = false, DiagnosticWritten = false, DiagnosticPath = null };
        }
        return result;
    }
}
