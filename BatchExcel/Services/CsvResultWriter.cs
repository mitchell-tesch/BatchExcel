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
    /// Index, Title, Status, Duration (ms), [output field columns...]
    /// </summary>
    public static void Write(string outFolder, BatchConfig config)
    {
        string csvPath = Path.Combine(outFolder, FileName);
        using var writer = new StreamWriter(csvPath);

        // Header row: identification columns + duration + output field column headers.
        // Every header is run through EscapeCsv so future field/sheet names that happen to contain
        // commas, quotes, or formula-leading chars cannot corrupt the CSV.
        var header = new List<string> { "Index", "Title", "Status", "Duration (ms)" };
        header.AddRange(config.OutputFields.Select(f => $"{f.Sheet}_{f.Range}"));
        writer.WriteLine(string.Join(",", header.Select(EscapeCsv)));

        // Data rows. Build raw cell strings first, then escape everything in one pass so we
        // never mix pre-escaped and unescaped values in the same row (a footgun if Status ever
        // grew to include user-supplied text).
        foreach (var run in config.Calculations)
        {
            string status;
            if (!run.Include)
                status = "Skipped";
            else if (run.Results == null)
                status = "Failed";
            else
                status = "Completed";

            // Duration is null for skipped/failed/never-executed runs — render as blank cell
            // rather than "0" so spreadsheet sorts/filters distinguish "didn't run" from
            // "ran in <1 ms".
            var duration = run.DurationMs?.ToString(CultureInfo.InvariantCulture) ?? "";

            var row = new List<string>(4 + config.OutputFields.Count)
            {
                (run.Index + 1).ToString(CultureInfo.InvariantCulture),
                run.Title,
                status,
                duration
            };

            if (run.Results != null)
            {
                foreach (var v in run.Results)
                    row.Add(FormatValue(v));
            }
            else
            {
                // Pad empty cells for output fields (Skipped or Failed)
                for (int i = 0; i < config.OutputFields.Count; i++)
                    row.Add("");
            }

            writer.WriteLine(string.Join(",", row.Select(EscapeCsv)));
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
        //
        // BUT: legitimate negative numbers (e.g. "-5.5") also start with '-'. Don't prefix
        // those — engineering calculations routinely produce negative outputs, and turning
        // them into text strings would corrupt downstream numeric analysis. Only neutralise
        // when the value is NOT a plain InvariantCulture number.
        bool startsWithFormulaChar = s[0] == '=' || s[0] == '+' || s[0] == '-' || s[0] == '@';
        bool needsFormulaPrefix = startsWithFormulaChar
            && !double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
        string value = needsFormulaPrefix ? "'" + s : s;

        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }
}

