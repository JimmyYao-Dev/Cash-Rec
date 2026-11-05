using System.Globalization;
using System.IO.Compression;

namespace CashRec;

public sealed record InputDefinition(string Key, string RelativePath);

public static class WorkspaceInputs
{
    public static readonly InputDefinition[] Definitions =
    [
        new("BankAU", "Bank/NAB AU.csv"), new("BankGL", "Bank/NAB GL.csv"),
        new("Transactions", "MT4/MT4/Transaction Summary Report.csv"),
        new("Equity", "MT4/Risk/Accounting Equity Report.csv"),
        new("BankBD", "PSP+BANK/BANK BD.xlsx"), new("PSP", "PSP+BANK/PSP Balance Summary.xlsx"),
        new("Rates", "Rate/VFSC Rate.xlsx")
    ];

    public static string DateDirectory(string workspace, DateTime date)
    {
        var root = Path.GetFullPath(workspace);
        var name = date.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture);
        if (DateTime.TryParseExact(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar)), "yyyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            return root;
        var candidates = new[] { Path.Combine(root, name), Path.Combine(root, "data", "risk_and_cash_to_process", name), Path.Combine(root, "risk_and_cash_to_process", name) };
        return candidates.FirstOrDefault(Directory.Exists) ?? candidates[0];
    }

    public static PreflightResult Inspect(string workspace, DateTime date)
    {
        var diagnostics = new List<Diagnostic>();
        var inputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(workspace))
            return new(date, "", inputs, "", [new("Error", "工作目录", "请选择业务工作目录")]);
        var root = Path.GetFullPath(workspace);
        var directory = DateDirectory(root, date);
        var folderName = Path.GetFileName(directory.TrimEnd(Path.DirectorySeparatorChar));
        if (DateTime.TryParseExact(folderName, "yyyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var folderDate) && folderDate.Date != date.Date)
            diagnostics.Add(new("Error", "日期", $"选中目录为 {folderName}，与业务日期 {date:yyyy.MM.dd} 不一致"));
        if (!Directory.Exists(directory)) diagnostics.Add(new("Error", "日期目录", $"缺少 {directory}"));
        foreach (var definition in Definitions)
        {
            var path = Path.Combine(directory, "Needed files", definition.RelativePath.Replace('/', Path.DirectorySeparatorChar));
            inputs[definition.Key] = path;
            CheckReadable(path, diagnostics);
        }
        var mapping = FindMapping(root, directory);
        if (mapping == "") diagnostics.Add(new("Error", "映射", "缺少 account_mapping.xlsx；请导入映射或放入工作区 config 目录"));
        else CheckReadable(mapping, diagnostics);
        if (!diagnostics.Any(d => d.Severity == "Error")) diagnostics.Add(new("Info", "准备检查", $"{date:yyyy.MM.dd}：7项输入及账户映射均可读取"));
        return new(date.Date, directory, inputs, mapping, diagnostics);
    }

    static void CheckReadable(string path, List<Diagnostic> diagnostics)
    {
        if (!File.Exists(path)) { diagnostics.Add(new("Error", Path.GetFileName(path), "缺少文件：" + path)); return; }
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (stream.Length == 0) throw new InvalidDataException("文件为空");
            if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
                if (zip.GetEntry("xl/workbook.xml") == null) throw new InvalidDataException("不是有效的 Excel 工作簿");
            }
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or InvalidDataException)
        { diagnostics.Add(new("Error", Path.GetFileName(path), "无法读取：" + error.Message)); }
    }

    static string FindMapping(string workspace, string dateDirectory)
    {
        var roots = new List<string> { workspace };
        var parent = new DirectoryInfo(dateDirectory);
        for (var i = 0; i < 5 && parent != null; i++, parent = parent.Parent) roots.Add(parent.FullName);
        return roots.Distinct(StringComparer.OrdinalIgnoreCase)
            .SelectMany(root => new[] { Path.Combine(root, "config", "account_mapping.xlsx"), Path.Combine(root, "db", "tmc", "account_mapping.xlsx") })
            .FirstOrDefault(File.Exists) ?? "";
    }

    public static string Import(string workspace, DateTime date, string source)
    {
        var name = Path.GetFileName(source);
        string target;
        if (name.Equals("account_mapping.xlsx", StringComparison.OrdinalIgnoreCase)) target = Path.Combine(Path.GetFullPath(workspace), "config", name);
        else
        {
            var definition = Definitions.FirstOrDefault(d => Path.GetFileName(d.RelativePath).Equals(name, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException("无法识别文件名：" + name);
            target = Path.Combine(DateDirectory(workspace, date), "Needed files", definition.RelativePath.Replace('/', Path.DirectorySeparatorChar));
        }
        if (Path.GetFullPath(source).Equals(Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase)) return target;
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        if (File.Exists(target)) File.Copy(target, target + ".before-import-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        File.Copy(source, target, true);
        return target;
    }
}
