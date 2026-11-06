namespace CashRec;

public sealed record MonthlyColumn(string Field, string Label, string Group = "");

public static class MonthlyLayout
{
    public const string Rules = "PRD1.6-Monthly-v1";
    public static readonly string[] Sheets = ["TMA", "TMG", "Total", "Break down"];
    public static readonly MonthlyColumn[] EntityColumns =
    [
        new("Date", ""), new("PSP DEP", "PSP DEP", "MT4"), new("PSP WITH", "PSP WITH", "MT4"),
        new("NET PSP", "NET PSP", "MT4"), new("NET BW", "NET BW", "MT4"), new("NET WS", "NET WS", "MT4"), new("", ""),
        new("Insto", "Insto", "BANK"), new("Retail", "Retail", "BANK"), new("Wholesale", "Wholesale", "BANK"),
        new("PSP", "PSP", "BANK"), new("PSP Pending", "PSP Pending", "BANK"), new("Settle", "Settle", "BANK"),
        new("SETTLE PNL", "SETTLE PNL", "BANK"), new("", ""),
        new("Insto DIFF", "Insto", "DIFF"), new("Insto Pending DIFF", "Insto Pending", "DIFF"),
        new("Retail DIFF", "Retail", "DIFF"), new("Retail Pending DIFF", "Retail Pending", "DIFF"),
        new("Wholesale DIFF", "Wholesale", "DIFF"), new("Wholesale Pending DIFF", "Wholesale Pending", "DIFF"),
        new("PSP DIFF", "PSP", "DIFF"), new("PSP Pending DIFF", "PSP Pending", "DIFF"), new("Settle DIFF", "Settle", "DIFF"), new("", ""),
    ];

    public static int HeaderRows(string sheet) => sheet == "Break down" ? 1 : 2;
    public static readonly MonthlyColumn[] BreakdownColumns =
    [
        new("Date", "TMGM Group"), new("Settle Bal", "Settle Bal"),
        new("Client Bank Move", "Client Bank Move (Inc lp&all pending)"), new("MT4", "MT4"), new("WS DEP", "WS DEP"),
        new("Diff", "Diff (C-D-E)"), new("To LP", "To LP"), new("TMA SETTLE PNL", "To settle pnl (TMA)"),
        new("TMG SETTLE PNL", "To Settle adj (TMG)"), new("To Settle WS", "To Settle (WS)"),
        new("Commission", "To Commission"), new("Bank Fee", "Bank Fee"), new("Bank fee-conversion", "Bank fee-conversion"),
        new("Time Diff", "Time Diff"), new("Error", "Error"), new("Check", "Check"),
    ];

    public static string Address(string sheet, string field, int row)
    {
        var columns = sheet == "Break down" ? BreakdownColumns : EntityColumns;
        var column = Array.FindIndex(columns, candidate => candidate.Field == field) + 1;
        if (column == 0) throw new InvalidDataException("月度缺少字段 " + field);
        return WorkbookWriter.CellAddress(column, row);
    }
}
