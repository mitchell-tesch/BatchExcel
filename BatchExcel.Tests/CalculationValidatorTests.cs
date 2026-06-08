using BatchExcel.Models;
using BatchExcel.Services;
using ClosedXML.Excel;

namespace BatchExcel.Tests;

public class CalculationValidatorTests : IDisposable
{
    private readonly string _tempPath;

    public CalculationValidatorTests()
    {
        _tempPath = Path.Combine(Path.GetTempPath(), $"test_template_{Guid.NewGuid()}.xlsx");
    }

    public void Dispose()
    {
        if (File.Exists(_tempPath))
            File.Delete(_tempPath);
        GC.SuppressFinalize(this);
    }

    private void CreateTestTemplate(string[] sheets, Dictionary<string, string>? namedRanges = null)
    {
        using var workbook = new XLWorkbook();
        foreach (var sheetName in sheets)
            workbook.Worksheets.Add(sheetName);

        if (namedRanges != null)
        {
            foreach (var nr in namedRanges)
                workbook.DefinedNames.Add(nr.Key, nr.Value);
        }

        workbook.SaveAs(_tempPath);
    }

    [Fact]
    public void Validate_ValidTemplate_DoesNotThrow()
    {
        CreateTestTemplate(
            new[] { "Sheet1", "Main" },
            new Dictionary<string, string> { ["JobNumber"] = "Sheet1!$A$1" });

        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "A1", 0) },
            OutputFields = { new FieldDefinition("Main", "B2", 0) },
            HeaderInputs = { ["JobNumber"] = "123" }
        };

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_EmptyConfig_DoesNotThrow()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig();

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_MissingSheet_ThrowsValidationException()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("MissingSheet", "A1", 0) }
        };

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("Missing sheet 'MissingSheet'", ex.Message);
        Assert.Contains("input field", ex.Message);
    }

    [Fact]
    public void Validate_MissingSheetForOutputField_ThrowsValidationException()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            OutputFields = { new FieldDefinition("Nope", "B2", 0) }
        };

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("Missing sheet 'Nope'", ex.Message);
        Assert.Contains("output field", ex.Message);
    }

    [Fact]
    public void Validate_HeaderInputs_AreNotEnforced()
    {
        // Header inputs are intentionally not validated — the writer ignores missing ones.
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            HeaderInputs = { ["MissingRange"] = "val" }
        };

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_A1Range_IsAccepted()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields =
            {
                new FieldDefinition("Sheet1", "A1", 0),
                new FieldDefinition("Sheet1", "$B$5", 1),
                new FieldDefinition("Sheet1", "AA100:AB200", 2),
                new FieldDefinition("Sheet1", "C10:D20", 3),
            }
        };

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_NamedRangeAsFieldRange_IsAccepted()
    {
        CreateTestTemplate(
            new[] { "Sheet1" },
            new Dictionary<string, string> { ["MyInput"] = "Sheet1!$A$1" });

        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "MyInput", 0) }
        };

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_UnknownNonA1Range_ThrowsValidationException()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "NotARangeOrName", 0) }
        };

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("'NotARangeOrName'", ex.Message);
        Assert.Contains("not a valid A1 reference", ex.Message);
    }

    [Fact]
    public void Validate_EmptyRange_ThrowsValidationException()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "", 0) }
        };

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("Empty range", ex.Message);
    }

    [Fact]
    public void Validate_FileInUse_CanStillRead()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "A1", 0) }
        };

        // Simulate file being open in Excel — exclusive write but shared read.
        using var fs = new FileStream(_tempPath, FileMode.Open, FileAccess.Read, FileShare.Read);

        CalculationValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_AggregatesMultipleErrors()
    {
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields =
            {
                new FieldDefinition("MissingSheet1", "A1", 0),
                new FieldDefinition("Sheet1", "BogusRange", 1),
            },
            OutputFields = { new FieldDefinition("MissingSheet2", "B2", 0) }
        };

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("Missing sheet 'MissingSheet1'", ex.Message);
        Assert.Contains("Missing sheet 'MissingSheet2'", ex.Message);
        Assert.Contains("'BogusRange'", ex.Message);
    }

    [Fact]
    public void Validate_CorruptFile_ThrowsValidationException()
    {
        // Write garbage bytes — not a valid .xlsx package.
        File.WriteAllBytes(_tempPath, new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04 });
        var config = new BatchConfig();

        var ex = Assert.Throws<ValidationException>(
            () => CalculationValidator.Validate(_tempPath, config));
        Assert.Contains("not a valid .xlsx file", ex.Message);
    }
}

