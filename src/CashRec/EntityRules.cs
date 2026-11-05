namespace CashRec;

/// <summary>Entity allocation for the two explicitly confirmed regulator fields.</summary>
public static class EntityRules
{
    public static string? RegulatorBranch(string value, out string exclusionReason)
    {
        var regulator = value.Trim();
        var asic = regulator.Contains("ASIC", StringComparison.OrdinalIgnoreCase);
        var vfsc = regulator.Contains("VFSC", StringComparison.OrdinalIgnoreCase);
        if (asic && vfsc)
        {
            exclusionReason = "监管歧义：同时包含 ASIC 与 VFSC";
            return null;
        }
        if (!asic && !vfsc)
        {
            exclusionReason = regulator.Length == 0 ? "监管为空" : "监管不在 TMA/TMG 范围";
            return null;
        }
        exclusionReason = "";
        return asic ? "TMA" : "TMG";
    }
}
