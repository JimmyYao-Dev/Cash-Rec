using System.Diagnostics;
using System.Text.Json;

namespace CashRec;

public sealed class MainForm : Form
{
    readonly TextBox workspaceBox = new() { Name = "workspaceBox", AccessibleName = "workspaceBox", Dock = DockStyle.Fill };
    readonly DateTimePicker businessDatePicker = new() { Name = "businessDatePicker", AccessibleName = "businessDatePicker", Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy.MM.dd", Width = 150 };
    readonly Button checkButton = new() { Name = "checkButton", AccessibleName = "checkButton", Text = "检查文件", AutoSize = true };
    readonly Button runButton = new() { Name = "runButton", AccessibleName = "runButton", Text = "生成基础报表", AutoSize = true };
    readonly Button importButton = new() { Name = "importButton", AccessibleName = "importButton", Text = "导入文件", AutoSize = true };
    readonly Button openButton = new() { Name = "openButton", AccessibleName = "openButton", Text = "打开结果位置", AutoSize = true, Enabled = false };
    readonly TextBox statusText = new() { Name = "statusText", AccessibleName = "statusText", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
    readonly Label resultLabel = new() { Name = "resultLabel", AutoSize = true, Text = "" };
    readonly ProgressBar progress = new() { Name = "progressBar", Dock = DockStyle.Top, Height = 5 };
    string? resultLocation;
    bool busy;
    readonly string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CashRec", "workspace.txt");

    public MainForm()
    {
        Text = "Cash Rec · 0.1.0";
        Name = "CashRecMain";
        Size = new Size(980, 680);
        MinimumSize = new Size(740, 520);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 10);
        BackColor = Color.FromArgb(246, 248, 250);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), RowCount = 7, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 5));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        layout.Controls.Add(new Label { Text = "Cash Rec", Font = new Font(Font.FontFamily, 22, FontStyle.Bold), AutoSize = true }, 0, 0);
        var workspaceRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        workspaceRow.Controls.Add(new Label { Text = "工作目录", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
        workspaceRow.Controls.Add(workspaceBox, 1, 0);
        var browse = new Button { Name = "browseButton", Text = "选择…", AutoSize = true };
        workspaceRow.Controls.Add(browse, 2, 0);
        layout.Controls.Add(workspaceRow, 0, 1);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(0, 8, 0, 0) };
        actions.Controls.AddRange([businessDatePicker, importButton, checkButton, runButton, openButton]);
        layout.Controls.Add(actions, 0, 2);
        layout.Controls.Add(progress, 0, 3);
        layout.Controls.Add(statusText, 0, 4);
        layout.Controls.Add(resultLabel, 0, 5);
        layout.Controls.Add(new Label { Text = "Windows x64  ·  v0.1.0", ForeColor = Color.DimGray, AutoSize = true }, 0, 6);
        Controls.Add(layout);
        if (File.Exists(settingsPath)) workspaceBox.Text = File.ReadAllText(settingsPath);
        statusText.Text = "选择业务工作目录和日期，然后检查文件。";
        browse.Click += (_, _) =>
        {
            using var dialog = new FolderBrowserDialog { Description = "选择业务工作目录", UseDescriptionForTitle = true };
            if (dialog.ShowDialog(this) == DialogResult.OK) workspaceBox.Text = dialog.SelectedPath;
        };
        checkButton.Click += (_, _) => CheckInputs();
        importButton.Click += (_, _) => ImportFiles();
        runButton.Click += async (_, _) => await Generate();
        openButton.Click += (_, _) =>
        {
            if (resultLocation == null) return;
            var launch = new ProcessStartInfo("explorer.exe") { UseShellExecute = true };
            launch.ArgumentList.Add(resultLocation);
            Process.Start(launch);
        };
    }

    bool ValidWorkspace()
    {
        if (string.IsNullOrWhiteSpace(workspaceBox.Text)) { statusText.Text = "请选择业务工作目录。"; return false; }
        try
        {
            var path = Path.GetFullPath(workspaceBox.Text.Trim());
            var appPath = Path.GetFullPath(AppContext.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if ((path.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar).StartsWith(appPath, StringComparison.OrdinalIgnoreCase))
            { statusText.Text = "请使用程序安装目录以外的业务工作目录。"; return false; }
            workspaceBox.Text = path;
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
            File.WriteAllText(settingsPath, path);
            return true;
        }
        catch (Exception error) { statusText.Text = "工作目录：" + error.Message; return false; }
    }

    void ShowDiagnostics(IEnumerable<Diagnostic> diagnostics) => statusText.Text = string.Join(Environment.NewLine,
        diagnostics.Select(d => $"[{d.Severity}] {d.Source}：{d.Message}"));

    void CheckInputs()
    {
        if (busy || !ValidWorkspace()) return;
        try
        {
            var check = WorkspaceInputs.Inspect(workspaceBox.Text, businessDatePicker.Value.Date);
            ShowDiagnostics(check.Diagnostics);
            var path = ReportRunner.DiagnosticPath(workspaceBox.Text, businessDatePicker.Value.Date);
            File.WriteAllText(path, JsonSerializer.Serialize(check, new JsonSerializerOptions { WriteIndented = true }));
            SetResult(Path.GetDirectoryName(path)!, check.CanRun ? "文件准备检查完成" : "存在缺项，请查看检查记录");
        }
        catch (Exception error) { statusText.Text = "准备检查：" + error.Message; }
    }

    void ImportFiles()
    {
        if (busy || !ValidWorkspace()) return;
        using var dialog = new OpenFileDialog { Multiselect = true, Filter = "业务文件|*.csv;*.xlsx", Title = "选择源文件或 account_mapping.xlsx" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var messages = new List<Diagnostic>();
        foreach (var file in dialog.FileNames)
        {
            try { messages.Add(new("Info", "导入", WorkspaceInputs.Import(workspaceBox.Text, businessDatePicker.Value.Date, file))); }
            catch (Exception error) { messages.Add(new("Error", Path.GetFileName(file), error.Message)); }
        }
        ShowDiagnostics(messages);
    }

    async Task Generate()
    {
        if (busy || !ValidWorkspace()) return;
        busy = true;
        checkButton.Enabled = runButton.Enabled = importButton.Enabled = false;
        businessDatePicker.Enabled = workspaceBox.Enabled = false;
        progress.Style = ProgressBarStyle.Marquee;
        resultLabel.Text = "处理中…";
        var report = new Progress<string>(message => statusText.AppendText(Environment.NewLine + message));
        try
        {
            statusText.Text = $"处理业务日期 {businessDatePicker.Value:yyyy.MM.dd}";
            var workspace = workspaceBox.Text;
            var date = businessDatePicker.Value.Date;
            var result = await Task.Run(() => ReportRunner.Run(workspace, date, message => ((IProgress<string>)report).Report(message)));
            ShowDiagnostics(result.Diagnostics);
            if (result.OutputPath != null) statusText.AppendText(Environment.NewLine + "输出：" + result.OutputPath);
            SetResult(Path.GetDirectoryName(result.OutputPath ?? result.DiagnosticPath)!, result.IsComplete ? "每日基础报表已生成" : result.OutputPath != null ? "已生成部分结果，请查看诊断记录" : "未生成报表，请处理诊断中的问题");
        }
        catch (Exception error) { statusText.Text = "处理失败：" + error.Message; resultLabel.Text = "处理失败"; }
        finally
        {
            busy = false;
            checkButton.Enabled = runButton.Enabled = importButton.Enabled = true;
            businessDatePicker.Enabled = workspaceBox.Enabled = true;
            progress.Style = ProgressBarStyle.Blocks;
        }
    }

    void SetResult(string directory, string message)
    {
        resultLocation = directory;
        openButton.Enabled = true;
        resultLabel.Text = message;
    }
}
