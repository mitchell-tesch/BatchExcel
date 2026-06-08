using BatchExcel.Models;
using BatchExcel.Services;

namespace BatchExcel.Tests;

public class CsvResultWriterTests : IDisposable
{
    private readonly string _tempDir;

    public CsvResultWriterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "BatchExcelTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { /* ignored */ }
    }

    private static BatchConfig MakeConfig(params BatchRun[] runs)
    {
        var cfg = new BatchConfig();
        cfg.OutputFields.Add(new FieldDefinition("Sheet1", "B1", 1));
        cfg.OutputFields.Add(new FieldDefinition("Sheet1", "C1", 2));
        cfg.Calculations.AddRange(runs);
        return cfg;
    }

    private string ReadCsv()
    {
        return File.ReadAllText(Path.Combine(_tempDir, "raw_output_fields.csv"));
    }

    [Fact]
    public void Write_ProducesHeaderRow()
    {
        var cfg = MakeConfig();
        CsvResultWriter.Write(_tempDir, cfg);

        string csv = ReadCsv();
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Index,Title,Status,Sheet1_B1,Sheet1_C1", lines[0].TrimEnd('\r'));
    }

    [Fact]
    public void Write_DistinguishesCompletedSkippedAndFailed()
    {
        var cfg = MakeConfig(
            new BatchRun { Index = 0, Include = true, Title = "OK", Results = new object?[] { 1.0, 2.0 } },
            new BatchRun { Index = 1, Include = false, Title = "Skip" },
            new BatchRun { Index = 2, Include = true, Title = "Fail", Results = null });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        Assert.Contains("1,OK,Completed,1,2", csv);
        Assert.Contains("2,Skip,Skipped,,", csv);
        Assert.Contains("3,Fail,Failed,,", csv);
    }

    [Fact]
    public void Write_EscapesCommasQuotesAndNewlines()
    {
        var cfg = MakeConfig(
            new BatchRun
            {
                Index = 0, Include = true, Title = "has, comma",
                Results = new object?[] { "she said \"hi\"", "line1\nline2" }
            });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        Assert.Contains("\"has, comma\"", csv);
        Assert.Contains("\"she said \"\"hi\"\"\"", csv);
        Assert.Contains("\"line1\nline2\"", csv);
    }

    [Theory]
    [InlineData("=SUM(A1:A10)")]
    [InlineData("+1+1")]
    [InlineData("-CMD")]
    [InlineData("@inject")]
    public void Write_NeutralisesFormulaInjection(string injected)
    {
        var cfg = MakeConfig(
            new BatchRun
            {
                Index = 0, Include = true, Title = "T",
                Results = new object?[] { injected, "safe" }
            });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        // The injected value must be prefixed (and possibly quoted) so Excel doesn't evaluate it.
        Assert.DoesNotContain("," + injected + ",", csv);
        Assert.True(csv.Contains(",'" + injected) || csv.Contains(",\"'" + injected));
    }

    [Fact]
    public void Write_UsesInvariantCultureForNumbers()
    {
        var cfg = MakeConfig(
            new BatchRun
            {
                Index = 0, Include = true, Title = "N",
                Results = new object?[] { 1234.56, 0.5 }
            });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        Assert.Contains(",1234.56,", csv);
        Assert.Contains(",0.5", csv);
    }

    [Theory]
    [InlineData(-5.5)]
    [InlineData(-0.001)]
    [InlineData(-100.0)]
    public void Write_NegativeNumbers_AreNotEscapedAsFormulas(double v)
    {
        // Critical: engineering calculations routinely produce negative outputs. Wrapping a
        // legitimate "-5.5" in a single-quote prefix would convert it to the text string
        // "'-5.5" in Excel, corrupting downstream numeric analysis. The formula-injection
        // neutraliser must skip cells that parse as InvariantCulture numbers.
        var cfg = MakeConfig(
            new BatchRun
            {
                Index = 0, Include = true, Title = "Neg",
                Results = new object?[] { v, 0.0 }
            });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        var rendered = v.ToString(System.Globalization.CultureInfo.InvariantCulture);
        Assert.Contains("," + rendered + ",", csv);
        Assert.DoesNotContain(",'" + rendered, csv);
    }

    [Theory]
    [InlineData("-CMD|/c calc")]   // negative-looking, but not a number → must be neutralised
    [InlineData("-not_a_number")]
    [InlineData("+abc")]
    [InlineData("=1+1")]
    [InlineData("@inject")]
    public void Write_NonNumericValuesStartingWithFormulaChars_AreStillNeutralised(string injection)
    {
        // Ensures the negative-number carve-out from EscapeCsv doesn't open a security hole:
        // values that LOOK like they start with - or + but aren't valid numbers must still be
        // prefixed.
        var cfg = MakeConfig(
            new BatchRun
            {
                Index = 0, Include = true, Title = "Inj",
                Results = new object?[] { injection, "safe" }
            });

        CsvResultWriter.Write(_tempDir, cfg);
        string csv = ReadCsv();

        Assert.True(csv.Contains(",'" + injection) || csv.Contains(",\"'" + injection),
            $"Expected '{injection}' to be neutralised; got: {csv}");
    }
}

