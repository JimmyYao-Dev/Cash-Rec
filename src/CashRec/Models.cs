namespace CashRec;

public sealed record Diagnostic(string Severity, string Source, string Message);
public sealed record FormulaCell(string Formula, decimal CalculatedValue);

public sealed class ReportTable(string name)
{
    public string Name { get; } = name;
    public List<string> Columns { get; } = [];
    public List<Dictionary<string, object?>> Rows { get; } = [];
    public List<Diagnostic> Diagnostics { get; } = [];
    public void AddRow(Dictionary<string, object?> row)
    {
        foreach (var key in row.Keys) if (!Columns.Contains(key)) Columns.Add(key);
        Rows.Add(row);
    }
}

public sealed class RunContext(DateTime businessDate, string dateDirectory,
    Dictionary<string, string> inputs, string mappingPath, DailyRates rates,
    Action<string>? progress = null)
{
    public DateTime BusinessDate { get; } = businessDate.Date;
    public string DateDirectory { get; } = dateDirectory;
    public Dictionary<string, string> Inputs { get; } = inputs;
    public string MappingPath { get; } = mappingPath;
    public DailyRates Rates { get; } = rates;
    public Action<string>? Progress { get; } = progress;
    internal PreviousBalanceSnapshot? PreviousBalances { get; set; }
    internal List<UnmatchedManualEntry> UnmatchedManual { get; } = [];
    public string SourcePath(string key) => Inputs.TryGetValue(key, out var path)
        ? path : throw new InvalidDataException($"缺少输入：{key}");
}

public interface IReportAdapter
{
    string Name { get; }
    ReportTable Process(RunContext context);
}

public interface IReportFinalizer
{
    void FinalizeReport(RunContext context, IList<ReportTable> tables, List<Diagnostic> diagnostics);
}

public sealed record PreflightResult(DateTime BusinessDate, string DateDirectory,
    Dictionary<string, string> Inputs, string MappingPath, List<Diagnostic> Diagnostics)
{
    public bool CanRun => !Diagnostics.Any(d => d.Severity == "Error");
}

public sealed record RunResult(string? OutputPath, string DiagnosticPath, bool IsComplete,
    List<Diagnostic> Diagnostics);

internal sealed record DailyRunArtifacts(RunResult Result, RunContext? Context, IList<ReportTable> Tables, PreparedFile? PreparedOutput = null);

internal sealed record UnmatchedManualEntry(DateTime? Date, string Sheet, string Key, string Field, decimal? Value, string NumberFormat, string Reason);

internal sealed record PreviousBalanceSnapshot(DateTime Date, string DateDirectory, string Source,
    Dictionary<string, Dictionary<string, decimal>> Values,
    Dictionary<string, List<Dictionary<string, object?>>> Tables);
