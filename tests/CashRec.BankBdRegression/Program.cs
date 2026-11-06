using System.IO.Compression;
using System.Xml.Linq;
using CashRec;
using ClosedXML.Excel;

try
{
    Run();
    foreach (var workspace in args)
    {
        var result = ReportRunner.Run(Path.GetFullPath(workspace), new DateTime(2026, 10, 25));
        var expectedError = Path.GetFileName(workspace) switch
        {
            "bankbd-missing-currency" => "Currency",
            "bankbd-uncached-amount" => "缓存",
            _ => null,
        };
        if (expectedError != null)
            Assert(!result.IsComplete && result.OutputPath == null
                && result.Diagnostics.Any(d => d.Source == "BANK BD" && d.Severity == "Error" && d.Message.Contains(expectedError)),
                workspace + ": full runner must reject malformed BANK BD input");
        else
        {
            Assert(result.IsComplete && result.OutputPath != null, string.Join(";", result.Diagnostics.Select(d => d.Message)));
            if (Path.GetFileName(workspace) == "template-bankbd")
                Assert(result.Diagnostics.Count(d => d.Source == "BANK BD" && d.Message.Contains("跳过 232 条")) == 1,
                    "Full fixture template count must be one aggregate diagnostic");
        }
        Console.WriteLine("PASS fixture: " + Path.GetFileName(workspace));
    }
}
catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }

static void Run()
{
    var directory = Path.Combine(Path.GetTempPath(), "cashrec-bankbd-" + Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(directory);
    var path = Path.Combine(directory, "BANK BD.xlsx");
    var context = new RunContext(new DateTime(2026, 10, 25), directory, new() { ["BankBD"] = path },
        "", new DailyRates(new() { ["USD"] = 1m }));

    Write([
        [null, null, null, null, null, "TMA", null],
        [null, null, null, null, null, "TMG", 0],
        [null, null, null, null, null, "TMA", "=0"],
        [0, "=0", null, null, null, "TMG", "=SUM(A5:B5)"],
        [null, null, "  ", " ", " ", "TMA", 0],
        [0, null, "Genuine zero transaction", "Adjustment", "USD", "TMA-Retail", 0],
        [null, null, null, null, "USD", "TMG-Other", 0],
    ], new() { ["G4"] = ("0", false), ["B5"] = ("0", false), ["G5"] = ("0", false) });
    var table = new BankBdTransactionAdapter().Process(context);
    Assert(table.Diagnostics.All(d => d.Severity != "Error"), "Cached numeric zero template formulas must be accepted");
    Assert(table.Rows.Count == 2 && table.Rows.All(r => (decimal)r["Amount"]! == 0m), "Real zero transactions with business fields must remain");
    Assert(table.Rows.Any(r => (string)r["Comment"]! == "Genuine zero transaction"), "Meaningful zero comment preserved");
    Assert(table.Diagnostics.Count(d => d.Message.Contains("跳过 5 条")) == 1, "Template skip count is aggregated");

    ExpectError([1, null, null, null, null, "TMA", 0], "Currency", "Nonzero Credit cannot be a template");
    ExpectError([null, -1, null, null, null, "TMG", 0], "Currency", "Nonzero Debit cannot be a template");
    ExpectError([null, null, null, null, null, "TMA", 1], "Currency", "Nonzero Amount cannot be a template");
    ExpectError([null, null, "Meaningful transaction", null, null, "TMA-Retail", 0], "Currency", "Missing Currency must never default to USD");
    ExpectError([null, null, null, null, null, "TMA", "=0"], "缓存", "Uncached Amount formula must not become zero");
    ExpectError([null, null, null, null, null, "TMA", "=0"], "缓存", "Formula text zero is not a numeric cache",
        new() { ["G2"] = ("0", true) });
    ExpectError(["=0", null, null, null, null, "TMG", 0], "缓存", "Uncached Credit cannot make a template");
    ExpectError([null, null, "=\"\"", null, null, "TMA", 0], "缓存", "Uncached business text cannot be considered blank");
    ExpectError([null, null, null, null, null, "TMA", "invalid"], "Amount", "Invalid Amount text cannot make a template");

    using (var book = new XLWorkbook())
    {
        var sheet = book.AddWorksheet("BANK BD");
        string[] headers = ["Comment", "Rational", "Currency", "Company", "Amount"];
        for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
        sheet.Cell(2, 4).Value = "TMA";
        sheet.Cell(2, 5).Value = 0;
        book.SaveAs(path);
    }
    table = new BankBdTransactionAdapter().Process(context);
    Assert(table.Rows.Count == 0 && table.Diagnostics.Any(d => d.Message.Contains("跳过 1 条")), "Optional Credit/Debit columns need not exist");
    Console.WriteLine("PASS: BANK BD template occupancy, cached numeric zeros, meaningful zero retention, optional inputs and strict invalid/cache errors.");

    void ExpectError(object?[] row, string message, string label, Dictionary<string, (string Value, bool Text)>? caches = null)
    {
        Write([row], caches);
        var result = new BankBdTransactionAdapter().Process(context);
        Assert(result.Rows.Count == 0 && result.Diagnostics.Any(d => d.Severity == "Error" && d.Message.Contains(message)), label);
        Assert(result.Diagnostics.All(d => !d.Message.Contains("空模板")), label + ": invalid row must not be counted as template");
    }

    void Write(object?[][] rows, Dictionary<string, (string Value, bool Text)>? caches = null)
    {
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("BANK BD");
            string[] headers = ["Credit", "Debit", "Comment", "Rational", "Currency", "Company", "Amount"];
            for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
            for (var r = 0; r < rows.Length; r++)
            for (var c = 0; c < rows[r].Length; c++)
            {
                var cell = sheet.Cell(r + 2, c + 1);
                switch (rows[r][c])
                {
                    case null: break;
                    case string text when text.StartsWith('='): cell.FormulaA1 = text[1..]; break;
                    case string text: cell.Value = text; break;
                    case int number: cell.Value = number; break;
                    default: throw new InvalidOperationException("Unsupported synthetic value");
                }
            }
            book.SaveAs(path);
        }
        if (caches == null) return;
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
        XDocument document;
        using (var stream = entry.Open()) document = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        foreach (var cell in document.Descendants(ns + "c"))
        {
            if (!caches.TryGetValue((string)cell.Attribute("r")!, out var cached)) continue;
            cell.SetAttributeValue("t", cached.Text ? "str" : null);
            cell.Element(ns + "v")?.Remove();
            cell.Add(new XElement(ns + "v", cached.Value));
        }
        entry.Delete();
        using var output = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
        document.Save(output);
    }
}

static void Assert(bool condition, string label) { if (!condition) throw new InvalidOperationException(label); }
