using System.Globalization;
using System.Text;

namespace CashRec;

internal static class UnmatchedManualReport
{
    internal static PreparedFile? Prepare(string directory, string id, IReadOnlyCollection<UnmatchedManualEntry> entries)
    {
        if (entries.Count == 0) return null;
        var destination = Path.Combine(directory, "UnmatchedManual_" + id + ".csv");
        var temporary = OutputTransaction.TemporaryFor(destination);
        using (var writer = new StreamWriter(temporary, false, new UTF8Encoding(true)))
        {
            writer.WriteLine("Date,Sheet,Key,Field,Value,NumberFormat,Reason");
            foreach (var entry in entries)
                writer.WriteLine(string.Join(",", new[] { entry.Date?.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture) ?? "", entry.Sheet,
                    entry.Key, entry.Field, entry.Value?.ToString(CultureInfo.InvariantCulture) ?? "", entry.NumberFormat, entry.Reason }
                    .Select(value => "\"" + value.Replace("\"", "\"\"") + "\"")));
        }
        return new(destination, temporary, Path.Combine(directory, "backups"));
    }
}
