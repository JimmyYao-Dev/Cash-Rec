using ClosedXML.Excel;

namespace CashRec;

public static class WorkbookWriter
{
    public static readonly string[] SheetNames = ["Equity", "Bank", "PSP", "TransactionExchange"];
    public static void Write(string path, IList<ReportTable> tables)
    {
        using var workbook = new XLWorkbook();
        workbook.CustomProperties.Add("CashRecRules", "PRD1.6-ASIC-VFSC");
        workbook.CustomProperties.Add("CashRecVersion", "0.1.0");
        foreach (var name in SheetNames)
        {
            var sheet = workbook.AddWorksheet(name);
            var data = tables.Where(t => t.Name == name).ToList();
            var columns = data.SelectMany(t => t.Columns).Distinct().ToList();
            for (var c = 0; c < columns.Count; c++) sheet.Cell(1, c + 1).Value = columns[c];
            var rowNumber = 2;
            foreach (var row in data.SelectMany(t => t.Rows))
            {
                for (var c = 0; c < columns.Count; c++)
                {
                    if (!row.TryGetValue(columns[c], out var value) || value == null) continue;
                    var cell = sheet.Cell(rowNumber, c + 1);
                    switch (value)
                    {
                        case FormulaCell formula: cell.FormulaA1 = formula.Formula; break;
                        case DateTime date: cell.Value = date; cell.Style.DateFormat.Format = "yyyy.mm.dd"; break;
                        case decimal number: cell.Value = (double)number; cell.Style.NumberFormat.Format = columns[c].Contains("Rate", StringComparison.OrdinalIgnoreCase) || columns[c].Contains("Units", StringComparison.OrdinalIgnoreCase) ? "0.########" : "#,##0.00;[Red](#,##0.00);0.00"; break;
                        case double number: cell.Value = number; break;
                        case int number: cell.Value = number; break;
                        case bool boolean: cell.Value = boolean; break;
                        default: cell.Value = value.ToString() ?? ""; break;
                    }
                    if (columns[c] is "DEP" or "PNL" || columns[c].Contains("USD", StringComparison.OrdinalIgnoreCase))
                        cell.Style.NumberFormat.Format = "$#,##0.00;[Red]($#,##0.00);$0.00";
                }
                rowNumber++;
            }
            if (columns.Count > 0)
            {
                var header = sheet.Range(1, 1, 1, columns.Count);
                header.Style.Font.Bold = true;
                header.Style.Font.FontColor = XLColor.White;
                header.Style.Fill.BackgroundColor = XLColor.FromHtml("183B56");
                sheet.SheetView.FreezeRows(1);
                sheet.Range(1, 1, Math.Max(1, rowNumber - 1), columns.Count).SetAutoFilter();
                foreach (var column in sheet.Columns(1, columns.Count)) column.Width = 22;
            }
        }
        workbook.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = true });
    }
}
