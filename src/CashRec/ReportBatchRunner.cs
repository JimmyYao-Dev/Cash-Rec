namespace CashRec;

public sealed record BatchDayResult(DateTime Date, string Status, bool Blocked, string? DailyOutput,
    string? MonthlyOutput, string? DiagnosticPath, List<Diagnostic> Diagnostics,
    TransactionStatus? Transaction = null, bool DiagnosticWritten = true, string? DiagnosticAttemptedPath = null);

public sealed class BatchRunResult(string workspace, DateTime start, DateTime end, string diagnosticPath)
{
    public bool Batch => true;
    public string Operation => "Generate";
    public string Workspace { get; } = workspace;
    public string ResultsDirectory => Workspace;
    public DateTime StartDate { get; } = start.Date;
    public DateTime EndDate { get; } = end.Date;
    public string? DiagnosticPath { get; internal set; } = diagnosticPath;
    public string DiagnosticAttemptedPath { get; } = diagnosticPath;
    public bool DiagnosticWritten { get; internal set; } = true;
    public string Status { get; internal set; } = "Failed";
    public bool Complete => Status == "Success";
    public List<BatchDayResult> Days { get; } = [];
    public List<Diagnostic> Diagnostics { get; } = [];
    public Dictionary<string, int> Counts => new[] { "Success", "Missing", "Failed", "Partial" }.ToDictionary(status => status, status => Days.Count(day => day.Status == status));
    public List<string> PublishedOutputs => Days.SelectMany(day => new[] { day.DailyOutput, day.MonthlyOutput }).OfType<string>().Distinct(StringComparer.OrdinalIgnoreCase).ToList();
}

/// <summary>Sequences independent day tasks; a failed date never substitutes adjacent-day input.</summary>
public static class ReportBatchRunner
{
    public static BatchRunResult Run(string workspace, DateTime start, DateTime end,
        Action<string>? progress = null, Action<BatchDayResult>? completed = null)
    {
        workspace = WorkspaceInputs.RangeWorkspace(workspace);
        var directory = Path.Combine(workspace, ".cashrec", "diagnostics");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, $"batch-{start:yyyy.MM.dd}-{end:yyyy.MM.dd}-{Guid.NewGuid():N}.json");
        var result = new BatchRunResult(workspace, start, end, path);
        if (start.Date > end.Date)
            result.Diagnostics.Add(new("Error", "日期范围", $"开始日期{start:yyyy.MM.dd}晚于结束日期{end:yyyy.MM.dd}；未执行任何日期", DiagnosticCode.InvalidDateRange));
        else
        {
            var total = (end.Date - start.Date).Days + 1;
            var count = 0;
            foreach (var date in Dates(start, end))
            {
                count++;
                var prefix = $"{date:yyyy.MM.dd} ({count}/{total})";
                progress?.Invoke(prefix);
                BatchDayResult day;
                try
                {
                    var task = ReportTaskRunner.Run(workspace, date, message => progress?.Invoke(prefix + " " + message));
                    var errors = task.Diagnostics.Where(diagnostic => diagnostic.Severity == "Error").ToList();
                    var missing = errors.Count > 0 && errors.All(diagnostic => diagnostic.Code is DiagnosticCode.MissingInput or DiagnosticCode.MissingPspDate);
                    var status = task.Complete ? "Success" : task.DailyOutput != null ? "Partial" : missing ? "Missing" : "Failed";
                    var blocked = task.Diagnostics.Any(diagnostic => diagnostic.Code is DiagnosticCode.MissingPriorBasis or DiagnosticCode.MissingPriorKey or DiagnosticCode.MissingPriorPending);
                    day = new(date, status, blocked, task.DailyOutput, task.MonthlyOutput, task.DiagnosticPath, task.Diagnostics,
                        task.Transaction, task.DiagnosticWritten, task.DiagnosticAttemptedPath);
                }
                catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException or InvalidOperationException or ArgumentException)
                {
                    day = new(date, "Failed", false, null, null, null, [new("Error", "报表任务", error.Message)], DiagnosticWritten: false);
                }
                result.Days.Add(day);
                completed?.Invoke(day);
            }
            result.Status = result.Days.All(day => day.Status == "Success") ? "Success"
                : result.PublishedOutputs.Count > 0 ? "Partial"
                : result.Days.All(day => day.Status == "Missing") ? "Missing" : "Failed";
        }
        try { DiagnosticRecords.Write(path, result); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            result.DiagnosticWritten = false; result.DiagnosticPath = null;
            result.Diagnostics.Add(new("Error", "范围记录", $"范围记录未保存：{path}；{error.Message}", DiagnosticCode.DiagnosticWriteFailed));
            result.Status = result.PublishedOutputs.Count > 0 ? "Partial" : "Failed";
        }
        return result;
    }

    public static IEnumerable<DateTime> Dates(DateTime start, DateTime end)
    {
        if (start.Date > end.Date) throw new ArgumentException($"开始日期{start:yyyy.MM.dd}晚于结束日期{end:yyyy.MM.dd}");
        for (var date = start.Date; ; date = date.AddDays(1))
        {
            yield return date;
            if (date == end.Date) break;
        }
    }
}
