using System.Globalization;
using System.IO;
using BatchExcel.Models;

namespace BatchExcel.Services;

/// <summary>
/// Writes batch run results to a CSV file in the output folder.
/// Includes run identification columns (Index, Title, Status) for easy correlation.
/// </summary>
public static class CsvResultWriter
{
    private const string FileName = "raw_output_fields.csv";

    /// <summary>
    /// Writes results to CSV. Output format:
    /// Index, Title, Status, [output field columns...]
    /// </summary>
    public static void Write(string outFolder, BatchConfig config)
    {
        string csvPath = Path.Combine(outFolder, FileName);
        using var writer = new StreamWriter(csvPath);

        // Header row: identification columns + output field column headers
        var header = new List<string> { "Index", "Title", "Status" };
        header.AddRange(config.OutputFields.Select(f => $"{f.Sheet}_{f.Range}"));
        writer.WriteLine(string.Join(",", header.Select(EscapeCsv)));

        // Data rows
        foreach (var run in config.Calculations)
        {
            string status;
            if (!run.Include)
                status = "Skipped";
            else if (run.Results == null)
                status = "Failed";
            else
                status = "Completed";

            var row = new List<string>
            {
                (run.Index + 1).ToString(CultureInfo.InvariantCulture),
                EscapeCsv(run.Title),
                status
            };

            if (run.Results != null)
            {
                foreach (var v in run.Results)
                {
                    row.Add(EscapeCsv(FormatValue(v)));
                }
            }
            else
            {
                // Pad empty cells for output fields (Skipped or Failed)
                for (int i = 0; i < config.OutputFields.Count; i++)
                    row.Add("");
            }

            writer.WriteLine(string.Join(",", row));
        }
    }

    private static string FormatValue(object? v)
    {
        if (v == null) return "";
        return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
    }

    private static string EscapeCsv(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;

        // Defuse CSV/formula injection: cells starting with =, +, -, @ are interpreted
        // as formulas by Excel/LibreOffice when the CSV is opened. Prefix with a single quote.
        bool needsFormulaPrefix = s.Length > 0 && (s[0] == '=' || s[0] == '+' || s[0] == '-' || s[0] == '@');
        string value = needsFormulaPrefix ? "'" + s : s;

        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }
}

