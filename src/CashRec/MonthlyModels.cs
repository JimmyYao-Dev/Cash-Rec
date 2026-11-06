namespace CashRec;

public sealed record MonthlyFormula(string Expression, decimal CalculatedValue);

public static class MonthlyStage
{
    public const string Current = "Breakdown";
}

public sealed class MonthlyReport(DateTime date)
{
    public DateTime Date { get; } = date.Date;
    public DateTime? PreviousDate { get; internal set; }
    public Dictionary<string, Dictionary<string, object?>> Rows { get; } = new(StringComparer.Ordinal)
    {
        ["TMA"] = new(), ["TMG"] = new(), ["Total"] = new(), ["Break down"] = new(),
    };
    public List<Diagnostic> Diagnostics { get; } = [];
    public bool BalancesComplete { get; internal set; }
    public bool FlowsComplete { get; internal set; }
    public bool BreakdownComplete { get; internal set; }
    public bool StageComplete => BalancesComplete && FlowsComplete && BreakdownComplete;
    public List<object> Provenance { get; } = [];
}

public sealed record MonthlyRunResult(string? OutputPath, string Stage, bool StageComplete,
    List<Diagnostic> Diagnostics);

public sealed record ReportTaskResult(string? DailyOutput, string? MonthlyOutput, string DiagnosticPath,
    string MonthlyStage, bool MonthlyStageComplete, bool Complete, List<Diagnostic> Diagnostics);
