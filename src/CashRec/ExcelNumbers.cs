using System.Globalization;
using ClosedXML.Excel;

namespace CashRec;

/// <summary>Financial values use raw input precision and only numeric formula caches.</summary>
public static class ExcelNumbers
{
    public static decimal Read(IXLCell cell, string source, bool allowBlankZero = false)
    {
        var value = cell.HasFormula ? cell.CachedValue : cell.Value;
        if (allowBlankZero && !cell.HasFormula && value.Type == XLDataType.Blank) return 0m;
        if (value.Type == XLDataType.Number)
            return DataValue.Decimal(value.GetNumber().ToString("R", CultureInfo.InvariantCulture), source);
        if (cell.HasFormula) throw new InvalidDataException(source + "：公式缺少有效数值缓存");
        return DataValue.Decimal(cell.GetString(), source);
    }
}
