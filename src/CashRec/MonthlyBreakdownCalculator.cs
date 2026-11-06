namespace CashRec;

/// <summary>Reconciles client balance changes to signed client flows.</summary>
public static class MonthlyBreakdownCalculator
{
    public static readonly string[] AutomaticFields = ["Settle Bal", "Client Bank Move", "MT4", "WS DEP", "Diff", "Error"];
    public static readonly string[] ClientDifferenceFields = ["Insto DIFF", "Retail DIFF", "Wholesale DIFF", "PSP DIFF", "PSP Pending DIFF"];

    public static void Apply(MonthlyReport report)
    {
        var total = report.Rows["Total"];
        var row = report.Rows["Break down"];
        row["Date"] = report.Date;
        row["Error"] = null;
        row["Settle Bal"] = new MonthlyFormula("[Total:Settle]", Required("Settle"));
        var mt4 = Required("NET PSP") + Required("NET BW");
        var wholesale = Required("NET WS");
        row["MT4"] = new MonthlyFormula("[Total:NET PSP]+[Total:NET BW]", mt4);
        row["WS DEP"] = new MonthlyFormula("[Total:NET WS]", wholesale);
        if (ClientDifferenceFields.Any(field => !total.TryGetValue(field, out var value) || value == null))
        {
            row["Client Bank Move"] = row["Diff"] = null;
            report.Diagnostics.Add(new("Incomplete", "资金差额", "客户余额五项 DIFF 尚不完整；Client Bank Move 和 Diff 留空，不将缺失变化当零"));
            report.BreakdownComplete = false;
            return;
        }
        var clientMove = ClientDifferenceFields.Sum(Required);
        row["Client Bank Move"] = new MonthlyFormula(string.Join("+", ClientDifferenceFields.Select(field => "[Total:" + field + "]")), clientMove);
        row["Diff"] = new MonthlyFormula("[Client Bank Move]-[MT4]-[WS DEP]", clientMove - mt4 - wholesale);
        report.BreakdownComplete = true;
        report.Provenance.Add(new
        {
            Kind = "Break down difference", Date = report.Date,
            ClientBankMove = ClientDifferenceFields,
            Mt4 = "Total NET PSP + NET BW", Wholesale = "Total NET WS", Difference = "Client Bank Move - MT4 - WS DEP",
            Manual = "To LP, Bank Fee, Bank fee-conversion, Time Diff remain manual; Error is blank",
        });
        return;

        decimal Required(string field) => MonthlyBalanceCalculator.Number(total.GetValueOrDefault(field), "Break down " + field);
    }
}
