using System.Text.Json;

namespace CashRec;

/// <summary>Publishes only complete JSON records; report publication is a separate outcome.</summary>
internal static class DiagnosticRecords
{
    internal static void Write(string path, object record)
    {
        var temporary = Path.Combine(Path.GetDirectoryName(path)!, ".CashRec-diagnostic-" + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(record, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temporary, path);
        }
        finally
        {
            try { if (File.Exists(temporary)) File.Delete(temporary); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException)
            { /* Preserve the write error; a leftover temporary file is never a published record. */ }
        }
    }
}
