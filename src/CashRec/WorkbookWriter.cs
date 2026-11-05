using ClosedXML.Excel;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace CashRec;

public static class WorkbookWriter
{
    public static readonly string[] SheetNames = ["Equity", "Bank", "PSP", "TransactionExchange"];
    public static void Write(string path, IList<ReportTable> tables)
    {
        using var workbook = new XLWorkbook();
        workbook.CustomProperties.Add("CashRecRules", ReportPolicy.EntityRules);
        workbook.CustomProperties.Add("CashRecValuation", ReportPolicy.Valuation);
        workbook.CustomProperties.Add("CashRecVersion", "0.1.0");
        if (tables.Any(table => table.Name == "Equity" && table.Columns.Contains("Previous Date")))
            workbook.CustomProperties.Add("CashRecManualDEP", "1");
        var cachedValues = new Dictionary<string, Dictionary<string, decimal>>();
        foreach (var name in SheetNames)
        {
            var sheet = workbook.AddWorksheet(name);
            var formulas = new Dictionary<string, decimal>();
            cachedValues["xl/worksheets/sheet" + workbook.Worksheets.Count + ".xml"] = formulas;
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
                        case FormulaCell formula:
                            cell.FormulaA1 = formula.Formula;
                            formulas[CellAddress(c + 1, rowNumber)] = formula.CalculatedValue;
                            break;
                        case DateTime date: cell.Value = date; cell.Style.DateFormat.Format = "yyyy.mm.dd"; break;
                        case decimal number: cell.Value = (double)number; cell.Style.NumberFormat.Format = columns[c].Contains("Rate", StringComparison.OrdinalIgnoreCase) || columns[c].Contains("Units", StringComparison.OrdinalIgnoreCase) ? "0.########" : "#,##0.00;[Red](#,##0.00);0.00"; break;
                        case double number: cell.Value = number; break;
                        case int number: cell.Value = number; break;
                        case bool boolean: cell.Value = boolean; break;
                        default: cell.Value = value.ToString() ?? ""; break;
                    }
                    if (columns[c] is "DEP" or "PNL" or "Movement" || columns[c].Contains("USD", StringComparison.OrdinalIgnoreCase))
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
        workbook.SaveAs(path, new ClosedXML.Excel.SaveOptions { EvaluateFormulasBeforeSaving = false });
        PersistFormulaValues(path, cachedValues);
    }

    static string CellAddress(int column, int row)
    {
        var letters = "";
        while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
        return letters + row.ToString(CultureInfo.InvariantCulture);
    }

    static void PersistFormulaValues(string path, Dictionary<string, Dictionary<string, decimal>> cachedValues)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        foreach (var (entryPath, values) in cachedValues.Where(pair => pair.Value.Count > 0))
        {
            var entry = archive.GetEntry(entryPath) ?? throw new InvalidDataException("缺少输出工作表：" + entryPath);
            XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            foreach (var cell in xml.Descendants(ns + "c"))
            {
                if (!values.TryGetValue((string?)cell.Attribute("r") ?? "", out var number)) continue;
                cell.Attribute("t")?.Remove();
                cell.SetElementValue(ns + "v", number.ToString(CultureInfo.InvariantCulture));
            }
            entry.Delete();
            using var output = archive.CreateEntry(entryPath).Open();
            xml.Save(output);
        }
        var workbookEntry = archive.GetEntry("xl/workbook.xml")!;
        XDocument workbookXml;
        using (var stream = workbookEntry.Open()) workbookXml = XDocument.Load(stream);
        var calculation = workbookXml.Root!.Element(ns + "calcPr");
        if (calculation == null) { calculation = new XElement(ns + "calcPr"); workbookXml.Root.Add(calculation); }
        calculation.SetAttributeValue("calcMode", "auto");
        calculation.SetAttributeValue("fullCalcOnLoad", "1");
        calculation.SetAttributeValue("forceFullCalc", "1");
        workbookEntry.Delete();
        using var workbookOutput = archive.CreateEntry("xl/workbook.xml").Open();
        workbookXml.Save(workbookOutput);
    }
}
