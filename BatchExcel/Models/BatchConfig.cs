using System.Globalization;

namespace BatchExcel.Models;

/// <summary>
/// Defines a cell field mapping in the calculation workbook.
/// </summary>
public record FieldDefinition(string Sheet, string Range, int ColumnOffset);

/// <summary>
/// Represents a single batch run with its input data and result storage.
/// </summary>
public class BatchRun
{
    private readonly string? _runId;

    /// <summary>Zero-based row position in the data table; used to write results back to the right row.</summary>
    public int Index { get; init; }

    /// <summary>The batcher sheet's Run ID (column A), or the 1-based row position if blank. IDs can have gaps.</summary>
    public string RunId
    {
        get => string.IsNullOrEmpty(_runId) ? (Index + 1).ToString(CultureInfo.InvariantCulture) : _runId;
        init => _runId = value;
    }

    public bool Include { get; init; }
    public string Title { get; init; } = "";
    public object?[] Data { get; init; } = [];
    public object?[]? Results { get; set; }

    /// <summary>
    /// Wall-clock duration of the calculation portion of the run in milliseconds, or null
    /// if the run was skipped / never executed. Populated by <see cref="Services.ExcelWorker"/>
    /// just before <see cref="Results"/> is assigned, so save-artifact time is excluded —
    /// the number reflects "Excel calc + macro + read" only, which is the useful figure for
    /// identifying slow runs in a large batch.
    /// </summary>
    public long? DurationMs { get; set; }

    /// <summary>Number of times an Excel process died while processing this run.</summary>
    public int ExcelCrashCount { get; set; }
}

/// <summary>
/// Complete batch configuration read from the batcher spreadsheet.
/// </summary>
public class BatchConfig
{
    // Calculation settings
    public string CalculationFile { get; set; } = "";
    public string CalculationMacrosRaw { get; set; } = "";
    public string CalculationHeaderSheet { get; set; } = "";

    // Header inputs (written to calculation header sheet)
    public Dictionary<string, object?> HeaderInputs { get; set; } = new();

    // Field definitions
    public List<FieldDefinition> InputFields { get; set; } = [];
    public List<FieldDefinition> OutputFields { get; set; } = [];
    public List<FieldDefinition> SkipFields { get; set; } = [];

    // Batch calculations
    public List<BatchRun> Calculations { get; set; } = [];

    // Derived properties
    public List<string> Macros => ParseCsvList(CalculationMacrosRaw);

    public int IncludedCalcCount => Calculations.Count(r => r.Include);

    private static List<string> ParseCsvList(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return [];
        return raw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
    }
}


