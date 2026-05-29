using BatchExcel.Services;
using ClosedXML.Excel;

namespace BatchExcel.Tests;

/// <summary>
/// End-to-end tests for BatchWorkbookReader that generate a real .xlsx fixture
/// matching the batcher layout (Main sheet, headers, data table from row 15).
/// </summary>
public class BatcherReaderTests : IDisposable
{
    private readonly string _tempDir;

    public BatcherReaderTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "BatchExcelTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { /* ignored */ }
    }

    /// <summary>
    /// Builds a minimal valid batcher workbook with the layout BatchWorkbookReader expects.
    /// </summary>
    private string CreateBatcherFixture(
        string calculationFile = "calculation.xlsx",
        string macros = "",
        params (string title, bool include, object?[] data)[] runs)
    {
        string path = Path.Combine(_tempDir, "batcher.xlsx");

        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Main");

        // Header cells (from BatchWorkbookReader.HeaderInRange)
        sheet.Cell(3, 2).Value = calculationFile;
        sheet.Cell(4, 2).Value = macros;
        sheet.Cell(5, 2).Value = "Header";
        sheet.Cell(7, 2).Value = "JOB-001";
        sheet.Cell(8, 2).Value = "Project X";
        sheet.Cell(9, 2).Value = "Alice";
        sheet.Cell(10, 2).Value = "Notes";

        // Data table header rows (DataStartRow = 15)
        // Row 15: type row.   Col B = Status header (skip), C = in, D = in, E = out
        // Row 16: sheet names
        // Row 17: ranges
        // Column B (Status) must be non-empty for every row so the reader's row-extent
        // walk (which scans down col B) reaches the last data row.
        sheet.Cell(15, 2).Value = "Status";
        sheet.Cell(15, 3).Value = "in";
        sheet.Cell(15, 4).Value = "in";
        sheet.Cell(15, 5).Value = "out";

        sheet.Cell(16, 2).Value = "Sheet";
        sheet.Cell(16, 3).Value = "Calc";
        sheet.Cell(16, 4).Value = "Calc";
        sheet.Cell(16, 5).Value = "Calc";

        sheet.Cell(17, 2).Value = "Range";
        sheet.Cell(17, 3).Value = "B2";
        sheet.Cell(17, 4).Value = "B3";
        sheet.Cell(17, 5).Value = "B4";

        // Data rows start at row 18 (DataStartRow + DataHeaderRowCount = 15 + 3).
        // Layout: col B = include status, col C = title (in), col D = numeric input (in),
        // col E = output (left blank in the fixture so we can verify WriteResults populates it).
        // The `data` parameter populates input cols D onwards (data[0] → col D).
        int r = 18;
        foreach (var (title, include, data) in runs)
        {
            sheet.Cell(r, 2).Value = include ? "Yes" : "No";
            sheet.Cell(r, 3).Value = title;
            // Only populate the numeric input column (D); ignore any extra entries to keep output (E) blank.
            int inputCount = Math.Min(data.Length, 1);
            for (int c = 0; c < inputCount; c++)
            {
                if (data[c] is null) continue;
                sheet.Cell(r, 4 + c).Value = XLCellValue.FromObject(data[c]);
            }
            r++;
        }

        // Reader requires at least one data row even if no runs were supplied; add a sentinel.
        if (runs.Length == 0)
        {
            sheet.Cell(18, 2).Value = "No";
            sheet.Cell(18, 3).Value = "_sentinel";
        }

        wb.SaveAs(path);
        return path;
    }

    [Fact]
    public void ReadConfig_ParsesHeaderFields()
    {
        var path = CreateBatcherFixture("calculation.xlsx", "MacroA, MacroB");

        var cfg = BatcherReader.ReadConfig(path);

        Assert.Equal("calculation.xlsx", cfg.CalculationFile);
        Assert.Equal("MacroA, MacroB", cfg.CalculationMacrosRaw);
        Assert.Equal(new[] { "MacroA", "MacroB" }, cfg.Macros);
        Assert.Equal("JOB-001", cfg.HeaderInputs["JobNumber"]);
        Assert.Equal("Project X", cfg.HeaderInputs["Project"]);
        Assert.Equal("Alice", cfg.HeaderInputs["Designer"]);
    }

    [Fact]
    public void ReadConfig_ParsesFieldDefinitions()
    {
        string path = CreateBatcherFixture(
            runs:
            [
                ("Run1", true, new object?[] { "Run1", 10.0, null })
            ]);

        var cfg = BatcherReader.ReadConfig(path);

        Assert.Equal(2, cfg.InputFields.Count);   // title col + 1 data col are 'in' in fixture
        Assert.Single(cfg.OutputFields);
        Assert.Equal("Calc", cfg.OutputFields[0].Sheet);
        Assert.Equal("B4", cfg.OutputFields[0].Range);
    }

    [Fact]
    public void ReadConfig_ParsesRunsWithIncludeAndTitle()
    {
        string path = CreateBatcherFixture(
            runs:
            [
                ("Alpha", true,  new object?[] { "Alpha", 1.0, null }),
                ("Beta",  false, new object?[] { "Beta",  2.0, null }),
                ("Gamma", true,  new object?[] { "Gamma", 3.0, null }),
            ]);

        var cfg = BatcherReader.ReadConfig(path);

        Assert.Equal(3, cfg.Calculations.Count);
        Assert.True(cfg.Calculations[0].Include);
        Assert.False(cfg.Calculations[1].Include);
        Assert.True(cfg.Calculations[2].Include);
        Assert.Equal("Alpha", cfg.Calculations[0].Title);
        Assert.Equal(2, cfg.IncludedCalcCount);
    }

    [Fact]
    public void ReadConfig_BlankTitle_GetsAutoGeneratedName()
    {
        string path = CreateBatcherFixture(
            runs:
            [
                ("", true, new object?[] { null, 1.0, null }),
            ]);

        var cfg = BatcherReader.ReadConfig(path);
        Assert.Equal("Run 1", cfg.Calculations[0].Title);
    }

    [Fact]
    public void WriteResults_RoundTripsOutputValues()
    {
        string path = CreateBatcherFixture(
            runs:
            [
                ("Run1", true, new object?[] { "Run1", 10.0, null }),
                ("Run2", true, new object?[] { "Run2", 20.0, null }),
            ]);

        var cfg = BatcherReader.ReadConfig(path);
        cfg.Calculations[0].Results = new object?[] { 42.5 };
        cfg.Calculations[1].Results = new object?[] { 99.0 };

        string outFolder = Path.Combine(_tempDir, "out");
        Directory.CreateDirectory(outFolder);
        bool ok = BatcherReader.WriteResults(path, cfg, outFolder);
        Assert.True(ok);

        // Re-read via ClosedXML to verify the output column (E) on data rows 18 and 19
        using var wb = new XLWorkbook(path);
        var sheet = wb.Worksheet("Main");
        Assert.Equal(42.5, sheet.Cell(18, 5).GetDouble());
        Assert.Equal(99.0, sheet.Cell(19, 5).GetDouble());

        // Copy in output folder should have the same values
        string copyPath = Path.Combine(outFolder, Path.GetFileName(path));
        Assert.True(File.Exists(copyPath));
        using var wbCopy = new XLWorkbook(copyPath);
        Assert.Equal(42.5, wbCopy.Worksheet("Main").Cell(18, 5).GetDouble());
    }

    [Fact]
    public void WriteResults_SkippedAndFailedRuns_AreLeftBlank()
    {
        string path = CreateBatcherFixture(
            runs:
            [
                ("Done",    true,  new object?[] { "Done", 1.0, null }),
                ("Skipped", false, new object?[] { "Skipped", 2.0, null }),
                ("Failed",  true,  new object?[] { "Failed", 3.0, null }),
            ]);

        var cfg = BatcherReader.ReadConfig(path);
        cfg.Calculations[0].Results = new object?[] { 100.0 };
        // Runs[1] is excluded; Runs[2] kept Results=null (failed)

        string outFolder = Path.Combine(_tempDir, "out");
        Directory.CreateDirectory(outFolder);
        BatcherReader.WriteResults(path, cfg, outFolder);

        using var wb = new XLWorkbook(path);
        var sheet = wb.Worksheet("Main");
        Assert.Equal(100.0, sheet.Cell(18, 5).GetDouble());
        Assert.True(sheet.Cell(19, 5).IsEmpty());
        Assert.True(sheet.Cell(20, 5).IsEmpty());
    }
}

