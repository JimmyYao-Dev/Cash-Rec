using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace CashRec;

/// <summary>Preserves dated cells/styles and rebuilds generated formulas after date reordering.</summary>
internal static class MonthlyHistory
{
    sealed record Cell(XLCellValue Value, string Formula, decimal? Cache, IXLStyle Style);
    sealed record Row(DateTime? Date, int OriginalRow, Cell[] Cells, Dictionary<string, Dictionary<int, DateTime>> OriginalDates);
    static readonly string[] ScalarFields = ["Insto", "Retail", "Wholesale", "PSP", "PSP Pending", "Settle", "SETTLE PNL", "PSP DEP", "PSP WITH", "NET BW", "NET WS"];

    internal static Dictionary<string, Dictionary<string, decimal>> Rewrite(XLWorkbook book, MonthlyReport current,
        DateTime? retainBaseline = null, XLWorkbook? additionalBaseline = null, DateTime? baselineDate = null)
    {
        using var styles = new XLWorkbook();
        var history = new Dictionary<string, Dictionary<DateTime, Row>>();
        var notes = new Dictionary<string, List<Row>>();
        var defaults = new Dictionary<string, Cell[]>();
        var originalDatesByRow = MonthlyLayout.Sheets.ToDictionary(name => name, _ => new Dictionary<int, DateTime>());
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.Worksheet(name);
            var styleSheet = styles.AddWorksheet(name);
            var columns = Columns(name);
            var header = MonthlyLayout.HeaderRows(name);
            var rows = new Dictionary<DateTime, Row>();
            notes[name] = [];
            defaults[name] = Capture(sheet.Row(header + 1), styleSheet.Row(1), columns.Length);
            foreach (var source in sheet.RowsUsed(XLCellsUsedOptions.Contents).Where(row => row.RowNumber() > header))
            {
                if (retainBaseline != null && source.Cell(1).HasFormula) continue;
                var date = MonthlyWorkbookWriter.CellDate(source.Cell(1));
                if (date != null) originalDatesByRow[name][source.RowNumber()] = date.Value;
                if (retainBaseline != null && date != retainBaseline) continue;
                var saved = new Row(date, source.RowNumber(), Capture(source, styleSheet.Row(source.RowNumber() + 1), columns.Length), originalDatesByRow);
                if (date == null)
                {
                    notes[name].Add(saved);
                    if (name == "Break down")
                        foreach (var field in MonthlySettlementCalculator.ManualFields)
                        {
                            var cell = saved.Cells[Array.FindIndex(columns, column => column.Field == field)];
                            if (cell.Value.Type == XLDataType.Blank && cell.Formula.Length == 0) continue;
                            current.UnmatchedManual.Add(new(null, name, "Row " + saved.OriginalRow, field, Number(cell, field),
                                cell.Style.NumberFormat.Format, "人工值没有有效业务日期，未分配给其他日期"));
                        }
                }
                else if (!rows.TryAdd(date.Value, saved)) throw new InvalidDataException(name + " 存在重复业务日期 " + date.Value.ToString("yyyy.MM.dd"));
            }
            history[name] = rows;
        }
        if (additionalBaseline != null)
        {
            var baselineDates = MonthlyLayout.Sheets.ToDictionary(name => name, name => additionalBaseline.Worksheet(name).RowsUsed(XLCellsUsedOptions.Contents)
                .Where(row => row.RowNumber() > MonthlyLayout.HeaderRows(name) && !row.Cell(1).HasFormula)
                .Select(row => (Number: row.RowNumber(), Date: MonthlyWorkbookWriter.CellDate(row.Cell(1))))
                .Where(pair => pair.Date != null).ToDictionary(pair => pair.Number, pair => pair.Date!.Value));
            foreach (var name in MonthlyLayout.Sheets)
            {
                var source = additionalBaseline.Worksheet(name).Row(baselineDates[name].Single(pair => pair.Value == baselineDate).Key);
                var saved = new Row(baselineDate, source.RowNumber(), Capture(source, styles.AddWorksheet(name + " basis").Row(1), Columns(name).Length), baselineDates);
                if (!history[name].TryAdd(baselineDate!.Value, saved)) throw new InvalidDataException(name + " 跨月基准日期重复");
            }
        }
        var originalDates = history["TMA"].Keys.Order().ToList();
        if (MonthlyLayout.Sheets.Any(name => !history[name].Keys.Order().SequenceEqual(originalDates)))
            throw new InvalidDataException("月度四表的业务日期集合不一致，已保留原文件，不能猜测人工值归属");
        var dates = originalDates.Append(current.Date).Distinct().Order().ToList();
        var rowNumbers = MonthlyLayout.Sheets.ToDictionary(name => name,
            name => dates.Select((date, index) => (date, number: MonthlyLayout.HeaderRows(name) + index + 1)).ToDictionary(pair => pair.date, pair => pair.number));
        var generated = dates.Where(date => date != current.Date).ToDictionary(date => date, date => Rebuild(date, history));
        generated[current.Date] = current;
        var caches = MonthlyLayout.Sheets.ToDictionary(name => name, _ => new Dictionary<string, decimal>());
        foreach (var name in MonthlyLayout.Sheets)
        {
            var sheet = book.Worksheet(name); var columns = Columns(name); var header = MonthlyLayout.HeaderRows(name);
            var last = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? header;
            if (retainBaseline != null)
            {
                var formattedEnd = sheet.LastRowUsed(XLCellsUsedOptions.All)?.RowNumber() ?? header;
                if (formattedEnd > header) sheet.Rows(header + 1, formattedEnd).Delete();
            }
            else if (last > header) sheet.Range(header + 1, 1, last, columns.Length).Clear(XLClearOptions.Contents);
            foreach (var date in dates)
            {
                var number = rowNumbers[name][date];
                var source = history[name].GetValueOrDefault(date);
                var template = source?.Cells ?? history[name].Values.Where(row => row.Date < date).OrderByDescending(row => row.Date).FirstOrDefault()?.Cells ?? defaults[name];
                for (var c = 0; c < columns.Length; c++)
                {
                    var cell = sheet.Cell(number, c + 1); cell.Style = template[c].Style;
                    if (source == null) continue;
                    var saved = source.Cells[c];
                    if (saved.Formula.Length == 0) cell.Value = saved.Value;
                    else
                    {
                        cell.FormulaA1 = Remap(saved.Formula, name, source!.OriginalDates, rowNumbers);
                        if (saved.Cache != null) caches[name][Address(c + 1, number)] = saved.Cache.Value;
                    }
                }
                var values = generated[date].Rows[name];
                for (var c = 0; c < columns.Length; c++)
                {
                    var field = columns[c].Field; var cell = sheet.Cell(number, c + 1);
                    var automatic = name == "Break down"
                        ? MonthlyBreakdownCalculator.AutomaticFields.Contains(field) || MonthlySettlementCalculator.AutomaticFields.Contains(field)
                        : MonthlyWorkbookWriter.IsAutomaticField(field);
                    if (automatic) { cell.Clear(XLClearOptions.Contents); caches[name].Remove(Address(c + 1, number)); }
                    if (!values.TryGetValue(field, out var value) || value == null) continue;
                    switch (value)
                    {
                        case DateTime day: cell.Value = day; if (source == null) cell.Style.DateFormat.Format = "mm-dd-yy"; break;
                        case decimal amount: cell.Value = (double)amount; break;
                        case MonthlyFormula formula:
                            cell.FormulaA1 = Resolve(formula.Expression, name, date, rowNumbers);
                            caches[name][Address(c + 1, number)] = formula.CalculatedValue;
                            break;
                    }
                }
            }
            var noteRow = header + dates.Count + 1;
            foreach (var note in notes[name])
            {
                for (var c = 0; c < columns.Length; c++)
                {
                    var cell = sheet.Cell(noteRow, c + 1); var saved = note.Cells[c]; cell.Style = saved.Style;
                    if (saved.Formula.Length == 0) cell.Value = saved.Value;
                    else { cell.FormulaA1 = Remap(saved.Formula, name, note.OriginalDates, rowNumbers); if (saved.Cache != null) caches[name][Address(c + 1, noteRow)] = saved.Cache.Value; }
                }
                noteRow++;
            }
        }
        return caches;
    }

    static MonthlyReport Rebuild(DateTime date, Dictionary<string, Dictionary<DateTime, Row>> history)
    {
        var report = new MonthlyReport(date);
        foreach (var name in new[] { "TMA", "TMG" })
        {
            var saved = history[name][date]; var values = report.Rows[name]; values["Date"] = date;
            foreach (var field in ScalarFields)
                if (Read(name, saved, field) is decimal number) values[field] = number;
            if (values.TryGetValue("PSP DEP", out var deposit) && values.TryGetValue("PSP WITH", out var withdrawal))
                values["NET PSP"] = new MonthlyFormula("[PSP DEP]+[PSP WITH]", (decimal)deposit! + (decimal)withdrawal!);
            foreach (var balance in MonthlyBalanceCalculator.BalanceFields)
            {
                var field = balance + " DIFF";
                var cell = Get(name, saved, field);
                if (cell.Formula.Length == 0) { if (Number(cell, name + " " + field) is decimal literal) values[field] = literal; continue; }
                var previous = Regex.Match(cell.Formula, @"^\$?[A-Za-z]+\$?\d+\s*-\s*\(?\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?)\s*\)?$", RegexOptions.CultureInvariant);
                if (!previous.Success) throw new InvalidDataException($"{name} {date:yyyy.MM.dd} {field} 历史公式无法按字段重绑，请重新生成该日期");
                var oldValue = DataValue.Decimal(previous.Groups[1].Value, field);
                var currentValue = MonthlyBalanceCalculator.Number(values.GetValueOrDefault(balance), balance);
                values[field] = new MonthlyFormula("[" + balance + "]-(" + oldValue.ToString(CultureInfo.InvariantCulture) + ")", currentValue - oldValue);
            }
        }
        var total = report.Rows["Total"]; total["Date"] = date;
        foreach (var field in ScalarFields.Concat(["NET PSP"]).Concat(MonthlyBalanceCalculator.BalanceFields.Select(field => field + " DIFF")).Distinct())
            if (report.Rows["TMA"].TryGetValue(field, out var tma) && report.Rows["TMG"].TryGetValue(field, out var tmg))
                total[field] = new MonthlyFormula("[TMA:" + field + "]+[TMG:" + field + "]", MonthlyBalanceCalculator.Number(tma, field) + MonthlyBalanceCalculator.Number(tmg, field));
        report.Rows["Break down"]["Date"] = date;
        if (new[] { "Settle", "NET PSP", "NET BW", "NET WS" }.All(total.ContainsKey))
        {
            MonthlyBreakdownCalculator.Apply(report);
            var commission = Read("Break down", history["Break down"][date], "Commission");
            if (commission != null && report.Rows["TMA"].ContainsKey("SETTLE PNL") && report.Rows["TMG"].ContainsKey("SETTLE PNL"))
            {
                var manual = MonthlySettlementCalculator.ManualFields.ToDictionary(field => field, field => Read("Break down", history["Break down"][date], field));
                MonthlySettlementCalculator.ApplyKnownCommission(report, commission.Value, manual);
            }
        }
        return report;
    }

    static Cell[] Capture(IXLRow source, IXLRow styles, int width)
    {
        var cells = new Cell[width];
        for (var c = 0; c < width; c++)
        {
            var cell = source.Cell(c + 1); styles.Cell(c + 1).Style = cell.Style;
            var value = cell.HasFormula ? cell.CachedValue : cell.Value;
            decimal? cache = cell.HasFormula && value.Type == XLDataType.Number ? DataValue.Decimal(value.GetNumber().ToString("R", CultureInfo.InvariantCulture), "历史公式缓存") : null;
            cells[c] = new(value, cell.HasFormula ? cell.FormulaA1 : "", cache, styles.Cell(c + 1).Style);
        }
        return cells;
    }
    static MonthlyColumn[] Columns(string sheet) => sheet == "Break down" ? MonthlyLayout.BreakdownColumns : MonthlyLayout.EntityColumns;
    static Cell Get(string sheet, Row row, string field) => row.Cells[Array.FindIndex(Columns(sheet), column => column.Field == field)];
    static decimal? Read(string sheet, Row row, string field) => Number(Get(sheet, row, field), $"{sheet} {row.Date:yyyy.MM.dd} {field}");
    static decimal? Number(Cell cell, string source)
    {
        if (cell.Formula.Length > 0) return cell.Cache ?? throw new InvalidDataException(source + " 公式缺少有效数值缓存");
        return cell.Value.Type == XLDataType.Blank ? null : cell.Value.Type == XLDataType.Number
            ? DataValue.Decimal(cell.Value.GetNumber().ToString("R", CultureInfo.InvariantCulture), source)
            : DataValue.Decimal(cell.Value.ToString(CultureInfo.InvariantCulture), source);
    }
    static string Address(int column, int row) => WorkbookWriter.CellAddress(column, row);
    static string Resolve(string expression, string sheet, DateTime date, Dictionary<string, Dictionary<DateTime, int>> rows) =>
        Regex.Replace(expression, @"\[([^\]]+)\]", match =>
        {
            var parts = match.Groups[1].Value.Split(':', 2); var target = parts.Length == 2 ? parts[0] : sheet;
            var address = MonthlyLayout.Address(target, parts[^1], rows[target][date]);
            return target == sheet ? address : "'" + target.Replace("'", "''") + "'!" + address;
        });
    static string Remap(string formula, string sheet, Dictionary<string, Dictionary<int, DateTime>> oldDates, Dictionary<string, Dictionary<DateTime, int>> rows)
    {
        var segments = Regex.Split(formula, "(\"(?:[^\"]|\"\")*\")");
        for (var i = 0; i < segments.Length; i += 2)
            segments[i] = Regex.Replace(segments[i], @"(?<![A-Za-z0-9_!])(?:(?<sheet>'(?:TMA|TMG|Total|Break down)'|TMA|TMG|Total)!)?(?<column>\$?[A-Za-z]{1,3})(?<absolute>\$?)(?<row>\d+)(?![A-Za-z0-9_(])", match =>
            {
                var requested = match.Groups["sheet"].Success ? match.Groups["sheet"].Value.Trim('\'') : sheet;
                var target = MonthlyLayout.Sheets.Single(name => name.Equals(requested, StringComparison.OrdinalIgnoreCase));
                if (!oldDates[target].TryGetValue(int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture), out var date)) return match.Value;
                if (!rows[target].ContainsKey(date)) throw new InvalidDataException("保留行公式引用已移除的历史日期，请先在Excel确认该人工值：" + target + " " + date.ToString("yyyy.MM.dd"));
                return (match.Groups["sheet"].Success ? match.Groups["sheet"].Value + "!" : "") + match.Groups["column"].Value + match.Groups["absolute"].Value + rows[target][date];
            }, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return string.Concat(segments);
    }
}
