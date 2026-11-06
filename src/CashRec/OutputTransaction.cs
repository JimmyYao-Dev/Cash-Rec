namespace CashRec;

internal sealed class PreparedFile(string destination, string temporaryPath, string backupDirectory) : IDisposable
{
    public string Destination { get; } = destination;
    public string TemporaryPath { get; } = temporaryPath;
    public string BackupDirectory { get; } = backupDirectory;
    public void Dispose() { if (File.Exists(TemporaryPath)) File.Delete(TemporaryPath); }
}

public sealed class TransactionStatus
{
    public string Id { get; } = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "-" + Guid.NewGuid().ToString("N");
    public string Status { get; internal set; } = "NotCommitted";
    public List<string> Backups { get; } = [];
    public List<string> CommittedOutputs { get; } = [];
    public List<string> RecoveryFiles { get; } = [];
    public List<string> CleanupWarnings { get; } = [];
}

internal static class OutputTransaction
{
    internal static FileStream Guard(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);

    public static void Commit(IReadOnlyList<PreparedFile> outputs, TransactionStatus status)
    {
        var originals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var stashed = new List<PreparedFile>();
        var committed = new List<PreparedFile>();
        var guards = new List<FileStream>();
        var targetGuards = new List<FileStream>();
        var phase = "检查报表目标";
        try
        {
            foreach (var output in outputs)
            {
                phase = "检查输出 " + output.Destination;
                if (Directory.Exists(output.Destination)) throw new IOException("输出目标是目录，无法替换文件：" + output.Destination);
                if (!targets.Add(output.Destination)) throw new InvalidDataException("任务存在重复输出目标");
                guards.Add(Guard(output.TemporaryPath));
                if (!File.Exists(output.Destination)) continue;
                if ((File.GetAttributes(output.Destination) & FileAttributes.ReadOnly) != 0) throw new UnauthorizedAccessException("输出是只读文件：" + output.Destination);
                targetGuards.Add(Guard(output.Destination));
                Directory.CreateDirectory(output.BackupDirectory);
                phase = "备份输出 " + output.Destination;
                var backup = Path.Combine(output.BackupDirectory, Path.GetFileNameWithoutExtension(output.Destination)
                    + "-" + status.Id + Path.GetExtension(output.Destination));
                File.Copy(output.Destination, backup);
                originals.Add(output.Destination, Path.Combine(Path.GetDirectoryName(output.Destination)!, ".CashRec-previous-" + status.Id + "-" + Path.GetFileName(output.Destination)));
                status.Backups.Add(backup);
            }
            // A guarded file can be renamed as a source, but Windows overwrite of
            // an open target needs write sharing. Never overwrite: retain originals
            // privately and move each protected candidate into an empty target.
            foreach (var output in outputs.Where(output => originals.ContainsKey(output.Destination)))
            {
                phase = "保留原结果 " + output.Destination;
                File.Move(output.Destination, originals[output.Destination]);
                stashed.Add(output);
            }
            foreach (var output in outputs)
            {
                phase = "替换输出 " + output.Destination;
                File.Move(output.TemporaryPath, output.Destination);
                committed.Add(output);
            }
            status.CommittedOutputs.AddRange(committed.Select(output => output.Destination));
            status.Status = "Committed";
            foreach (var output in stashed) Clean(originals[output.Destination]);
        }
        catch (Exception failure)
        {
            status.CommittedOutputs.Clear();
            var rollbackErrors = new List<string>();
            var discarded = new List<string>();
            foreach (var output in committed.AsEnumerable().Reverse())
            {
                try
                {
                    var discard = output.Destination + ".discard-" + status.Id;
                    File.Move(output.Destination, discard);
                    discarded.Add(discard);
                }
                catch (Exception error) { rollbackErrors.Add(output.Destination + "：" + error.Message); }
            }
            foreach (var output in stashed.AsEnumerable().Reverse())
                try { File.Move(originals[output.Destination], output.Destination); }
                catch (Exception error) { rollbackErrors.Add(output.Destination + "：" + error.Message); }
            foreach (var discard in discarded) Clean(discard);
            status.Status = rollbackErrors.Count > 0 ? "RollbackFailed" : committed.Count > 0 || stashed.Count > 0 ? "RolledBack" : "NotCommitted";
            if (rollbackErrors.Count > 0) status.RecoveryFiles.AddRange(originals.Values.Where(File.Exists));
            var explanation = rollbackErrors.Count == 0 ? "未发布新结果，原结果已保留" : "部分恢复失败，请使用备份恢复：" + string.Join("；", rollbackErrors);
            throw new IOException(phase + " 失败：" + failure.Message + "；" + explanation, failure);
        }
        finally { foreach (var guard in guards) guard.Dispose(); foreach (var guard in targetGuards) guard.Dispose(); }

        void Clean(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception error) { status.CleanupWarnings.Add(path + "：" + error.Message); }
        }
    }

    internal static string TemporaryFor(string destination) => Path.Combine(Path.GetDirectoryName(destination)!,
        ".CashRec-" + Guid.NewGuid().ToString("N") + Path.GetExtension(destination));
}
