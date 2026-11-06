using System.Diagnostics;
using System.Text.Json;

namespace CashRec;

public sealed class MainForm : Form
{
    readonly TextBox workspaceBox = new() { Name = "workspaceBox", AccessibleName = "workspaceBox", Dock = DockStyle.Fill };
    readonly DateTimePicker businessDatePicker = new() { Name = "businessDatePicker", AccessibleName = "businessDatePicker", Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy.MM.dd", Width = 150 };
    readonly CheckBox rangeCheckBox = new() { Name = "rangeCheckBox", AccessibleName = "rangeCheckBox", Text = "日期范围", AutoSize = true, Padding = new Padding(0, 4, 0, 0) };
    readonly DateTimePicker endDatePicker = new() { Name = "endDatePicker", AccessibleName = "endDatePicker", Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy.MM.dd", Width = 150, Visible = false };
    readonly Label toLabel = new() { Text = "至", AutoSize = true, Padding = new Padding(0, 5, 0, 0), Visible = false };
    readonly ToolTip inputHelp = new();
    readonly Button checkButton = new() { Name = "checkButton", AccessibleName = "checkButton", Text = "检查文件", AutoSize = true };
    readonly Button runButton = new() { Name = "runButton", AccessibleName = "runButton", Text = "生成报表", AutoSize = true };
    readonly Button importButton = new() { Name = "importButton", AccessibleName = "importButton", Text = "导入文件", AutoSize = true };
    readonly Button openButton = new() { Name = "openButton", AccessibleName = "openButton", Text = "打开结果位置", AutoSize = true, Enabled = false };
    readonly Button browseButton = new() { Name = "browseButton", Text = "选择…", AutoSize = true };
    readonly TextBox statusText = new() { Name = "statusText", AccessibleName = "statusText", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
    readonly Label resultLabel = new() { Name = "resultLabel", AutoSize = true, Text = "" };
    readonly ProgressBar progress = new() { Name = "progressBar", Dock = DockStyle.Top, Height = 5 };
    string? resultLocation;
    bool busy;
    readonly string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CashRec", "workspace.txt");

    public MainForm()
    {
        Text = "Cash Rec · " + AppVersion.Current;
        Name = "CashRecMain";
        Size = new Size(980, 680);
        MinimumSize = new Size(740, 520);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 10);
        BackColor = Color.FromArgb(246, 248, 250);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), RowCount = 6, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 5));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        layout.Controls.Add(new Label { Text = "Cash Rec", Font = new Font(Font.FontFamily, 22, FontStyle.Bold), AutoSize = true }, 0, 0);
        var workspaceRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        workspaceRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        workspaceRow.Controls.Add(new Label { Text = "工作目录", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
        workspaceRow.Controls.Add(workspaceBox, 1, 0);
        workspaceRow.Controls.Add(browseButton, 2, 0);
        layout.Controls.Add(workspaceRow, 0, 1);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Padding = new Padding(0, 8, 0, 8) };
        actions.Controls.AddRange([businessDatePicker, rangeCheckBox, toLabel, endDatePicker, importButton, checkButton, runButton, openButton]);
        layout.Controls.Add(actions, 0, 2);
        layout.Controls.Add(progress, 0, 3);
        layout.Controls.Add(statusText, 0, 4);
        layout.Controls.Add(resultLabel, 0, 5);
        Controls.Add(layout);
        if (File.Exists(settingsPath)) workspaceBox.Text = File.ReadAllText(settingsPath);
        statusText.Text = "选择业务工作目录和日期，然后检查文件。";
        browseButton.Click += (_, _) =>
        {
            using var dialog = new FolderBrowserDialog { Description = "选择业务工作目录", UseDescriptionForTitle = true };
            if (dialog.ShowDialog(this) == DialogResult.OK) workspaceBox.Text = dialog.SelectedPath;
        };
        inputHelp.SetToolTip(importButton, "导入按单日进行；日期范围使用各日期目录中已准备的文件。");
        inputHelp.SetToolTip(rangeCheckBox, "范围使用各日期目录中已准备的文件；需要导入时切回单日。");
        rangeCheckBox.CheckedChanged += (_, _) =>
        {
            if (rangeCheckBox.Checked) endDatePicker.Value = businessDatePicker.Value.Date;
            toLabel.Visible = endDatePicker.Visible = rangeCheckBox.Checked;
            importButton.Enabled = !busy && !rangeCheckBox.Checked;
        };
        checkButton.Click += async (_, _) => await CheckInputs();
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

    async Task CheckInputs()
    {
        if (busy || !ValidWorkspace()) return;
        SetBusy(true);
        try
        {
            if (rangeCheckBox.Checked)
            {
                var workspace = WorkspaceInputs.RangeWorkspace(workspaceBox.Text);
                var start = businessDatePicker.Value.Date; var end = endDatePicker.Value.Date;
                var checks = await Task.Run(() => ReportBatchRunner.Dates(start, end).Select(date => WorkspaceInputs.Inspect(workspace, date)).ToList());
                statusText.Text = "范围工作目录：" + workspace + Environment.NewLine + string.Join(Environment.NewLine,
                    checks.SelectMany(check => new[] { $"{check.BusinessDate:yyyy.MM.dd}：{(check.CanRun ? "文件齐备" : "存在问题")}" }
                        .Concat(check.Diagnostics.Select(diagnostic => $"[{diagnostic.Severity}] {diagnostic.Source}：{diagnostic.Message}"))));
                var rangePath = ReportRunner.DiagnosticPath(workspace, start);
                File.WriteAllText(rangePath, JsonSerializer.Serialize(new { Batch = true, Operation = "Check", StartDate = start, EndDate = end, Days = checks }, new JsonSerializerOptions { WriteIndented = true }));
                SetResult(workspace, checks.All(check => check.CanRun) ? "范围文件检查完成" : "部分日期存在问题，请查看检查记录");
                return;
            }
            var check = WorkspaceInputs.Inspect(workspaceBox.Text, businessDatePicker.Value.Date);
            ShowDiagnostics(check.Diagnostics);
            var path = ReportRunner.DiagnosticPath(workspaceBox.Text, businessDatePicker.Value.Date);
            File.WriteAllText(path, JsonSerializer.Serialize(check, new JsonSerializerOptions { WriteIndented = true }));
            SetResult(Path.GetDirectoryName(path)!, check.CanRun ? "文件准备检查完成" : "存在缺项，请查看检查记录");
        }
        catch (Exception error) { statusText.Text = "准备检查：" + error.Message; }
        finally { SetBusy(false); }
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
        SetBusy(true);
        resultLabel.Text = "处理中…";
        var report = new Progress<string>(message => statusText.AppendText(Environment.NewLine + message));
        try
        {
            statusText.Text = $"处理业务日期 {businessDatePicker.Value:yyyy.MM.dd}";
            var workspace = workspaceBox.Text;
            var date = businessDatePicker.Value.Date;
            if (rangeCheckBox.Checked)
            {
                var end = endDatePicker.Value.Date;
                statusText.Text = $"处理日期范围 {date:yyyy.MM.dd} 至 {end:yyyy.MM.dd}";
                var completed = new Progress<BatchDayResult>(day => statusText.AppendText(Environment.NewLine + DayText(day)));
                var batch = await Task.Run(() => ReportBatchRunner.Run(workspace, date, end,
                    message => ((IProgress<string>)report).Report(message), day => ((IProgress<BatchDayResult>)completed).Report(day)));
                statusText.Text = "范围工作目录：" + batch.Workspace + Environment.NewLine
                    + string.Join(Environment.NewLine, batch.Days.Select(DayText)) + Environment.NewLine
                    + string.Join(Environment.NewLine, batch.Diagnostics.Select(diagnostic => $"[{diagnostic.Severity}] {diagnostic.Source}：{diagnostic.Message}"))
                    + Environment.NewLine + "范围记录：" + batch.DiagnosticPath;
                var counts = batch.Counts;
                SetResult(batch.ResultsDirectory, batch.Diagnostics.Any(diagnostic => diagnostic.Code == DiagnosticCode.InvalidDateRange)
                    ? "日期范围无效，未执行任务"
                    : $"范围{(batch.Complete ? "完成" : StatusText(batch.Status))}；成功{counts["Success"]}，缺项{counts["Missing"]}，失败{counts["Failed"]}，部分完成{counts["Partial"]}");
                return;
            }
            var result = await Task.Run(() => ReportTaskRunner.Run(workspace, date, message => ((IProgress<string>)report).Report(message)));
            ShowDiagnostics(result.Diagnostics);
            if (result.DailyOutput != null) statusText.AppendText(Environment.NewLine + "每日基础表：" + result.DailyOutput);
            if (result.MonthlyOutput != null) statusText.AppendText(Environment.NewLine + "月度报表：" + result.MonthlyOutput);
            SetResult(Path.GetDirectoryName(result.MonthlyOutput ?? result.DailyOutput ?? result.DiagnosticPath)!,
                result.Complete ? "报表已生成；请在 Excel 完成人工核对"
                    : result.DailyOutput != null ? "结果不完整，请查看诊断记录" : "未生成报表，请处理诊断中的问题");
        }
        catch (Exception error) { statusText.Text = "处理失败：" + error.Message; resultLabel.Text = "处理失败"; }
        finally
        {
            SetBusy(false);
        }
    }

    void SetBusy(bool value)
    {
        busy = value;
        checkButton.Enabled = runButton.Enabled = !value;
        importButton.Enabled = !value && !rangeCheckBox.Checked;
        businessDatePicker.Enabled = endDatePicker.Enabled = rangeCheckBox.Enabled = workspaceBox.Enabled = browseButton.Enabled = !value;
        openButton.Enabled = !value && resultLocation != null;
        progress.Style = value ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
    }

    static string StatusText(string status) => status switch { "Success" => "成功", "Missing" => "缺项", "Partial" => "部分完成", _ => "失败" };
    static string DayText(BatchDayResult day) => $"{day.Date:yyyy.MM.dd}：{StatusText(day.Status)}{(day.Blocked ? "（上期基准受阻）" : "")}" + Environment.NewLine
        + string.Join(Environment.NewLine, day.Diagnostics.Select(diagnostic => $"[{diagnostic.Severity}] {diagnostic.Source}：{diagnostic.Message}"))
        + (day.DailyOutput == null ? "" : Environment.NewLine + "每日基础表：" + day.DailyOutput)
        + (day.MonthlyOutput == null ? "" : Environment.NewLine + "月度报表：" + day.MonthlyOutput)
        + (day.DiagnosticPath.Length == 0 ? "" : Environment.NewLine + "日期记录：" + day.DiagnosticPath);

    void SetResult(string directory, string message)
    {
        resultLocation = directory;
        openButton.Enabled = true;
        resultLabel.Text = message;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) inputHelp.Dispose();
        base.Dispose(disposing);
    }
}
