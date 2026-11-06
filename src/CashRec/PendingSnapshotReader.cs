using System.Globalization;
using System.Security.Cryptography;
using ClosedXML.Excel;

namespace CashRec;

public sealed record PendingValue(decimal PspSettled, decimal BankSettled, decimal Pending, int SourceRow);
public sealed record PendingSnapshot(DateTime Date, string SourcePath, string Sha256,
    IReadOnlyDictionary<string, PendingValue> Entities);

/// <summary>Consumes only the entity summary from the prepared day's cumulative snapshot.</summary>
public static class PendingSnapshotReader
{
    static readonly string[] Fields = ["Branch", "TOTAL PSP SETTLE USD", "TOTAL BANK SETTLE USD", "Total PSP On the way USD"];

    public static PendingSnapshot Read(string path, DateTime date)
    {
        var expectedDirectory = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)!, "..", ".."));
        if (!DateTime.TryParseExact(Path.GetFileName(expectedDirectory), "yyyy.MM.dd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var folderDate) || folderDate.Date != date.Date)
            throw new InvalidDataException($"PSP Pending：快照必须位于业务日期 {date:yyyy.MM.dd} 的 Needed files/PSP+BANK 目录");
        using var book = new XLWorkbook(path);
        var suppliedDate = book.CustomProperties.SingleOrDefault(property => property.Name == "CashRecSnapshotDate");
        if (suppliedDate != null && (!DateTime.TryParseExact(suppliedDate.Value.ToString(), "yyyy.MM.dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var declared) || declared.Date != date.Date))
            throw new InvalidDataException($"PSP Pending：快照声明日期与业务日期 {date:yyyy.MM.dd} 不一致");
        var sheet = book.Worksheets.SingleOrDefault(sheet => sheet.Name.Equals("PSP On The Way", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException("BANK BD：缺少 PSP On The Way 工作表");
        Dictionary<string, int>? headers = null;
        var headerRow = 0;
        foreach (var row in sheet.RowsUsed().Take(30))
        {
            var candidates = row.CellsUsed().Where(cell => !cell.HasFormula)
                .GroupBy(cell => DataValue.Normalize(cell.GetString()))
                .ToDictionary(group => group.Key, group => group.Select(cell => cell.Address.ColumnNumber).ToList());
            if (!Fields.All(field => candidates.TryGetValue(DataValue.Normalize(field), out var columns) && columns.Count == 1)) continue;
            if (headers != null) throw new InvalidDataException("PSP On The Way：汇总表头重复，无法确定快照区域");
            headers = Fields.ToDictionary(field => field, field => candidates[DataValue.Normalize(field)].Single());
            headerRow = row.RowNumber();
        }
        if (headers == null) throw new InvalidDataException("PSP On The Way：缺少唯一的 Branch、TOTAL PSP SETTLE USD、TOTAL BANK SETTLE USD、Total PSP On the way USD 汇总表头");
        var entities = new Dictionary<string, PendingValue>(StringComparer.Ordinal);
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow))
        {
            var branchCell = row.Cell(headers["Branch"]);
            var value = branchCell.HasFormula ? branchCell.CachedValue : branchCell.Value;
            if (branchCell.HasFormula && value.Type is XLDataType.Blank or XLDataType.Error)
                throw new InvalidDataException($"PSP On The Way!{branchCell.Address}：Branch 公式缺少有效缓存");
            var branch = value.ToString(CultureInfo.InvariantCulture).Trim().ToUpperInvariant();
            if (branch is not ("TMA" or "TMG")) continue;
            var source = $"{Path.GetFileName(path)} / PSP On The Way 第 {row.RowNumber()} 行";
            var psp = ExcelNumbers.Read(row.Cell(headers[Fields[1]]), source + " R / " + Fields[1]);
            var bank = ExcelNumbers.Read(row.Cell(headers[Fields[2]]), source + " S / " + Fields[2]);
            var pending = ExcelNumbers.Read(row.Cell(headers[Fields[3]]), source + " T / " + Fields[3]);
            // Excel stores IEEE doubles. Accept either exact decimal agreement or the exact
            // original floating subtraction; no business rounding or tolerance is introduced.
            if (psp - bank != pending && (double)psp - (double)bank != (double)pending)
                throw new InvalidDataException(source + "：T 与 R−S 不一致，无法使用在途结算快照");
            if (!entities.TryAdd(branch, new(psp, bank, pending, row.RowNumber())))
                throw new InvalidDataException(source + "：" + branch + " 汇总行重复，不能任选或累计");
        }
        foreach (var branch in new[] { "TMA", "TMG" })
            if (!entities.ContainsKey(branch)) throw new InvalidDataException("PSP On The Way：缺少 " + branch + " 汇总行，不当零");
        using var input = File.OpenRead(path);
        return new(date.Date, path, Convert.ToHexString(SHA256.HashData(input)), entities);
    }
}
