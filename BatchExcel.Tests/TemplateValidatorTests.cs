using BatchExcel.Models;
using BatchExcel.Services;
using ClosedXML.Excel;

namespace BatchExcel.Tests;

public class TemplateValidatorTests : IDisposable
{
    private readonly string _tempPath;

    public TemplateValidatorTests()
    {
        _tempPath = Path.Combine(Path.GetTempPath(), $"test_template_{Guid.NewGuid()}.xlsx");
    }

    public void Dispose()
    {
        if (File.Exists(_tempPath))
            File.Delete(_tempPath);
    }

    private void CreateTestTemplate(string[] sheets, Dictionary<string, string>? namedRanges = null)
    {
        using var workbook = new XLWorkbook();
        foreach (var sheetName in sheets)
        {
            workbook.Worksheets.Add(sheetName);
        }

        if (namedRanges != null)
        {
            foreach (var nr in namedRanges)
            {
                // Simple named range to a cell in the first sheet
                workbook.DefinedNames.Add(nr.Key, $"{sheets[0]}!$A$1");
            }
        }

        workbook.SaveAs(_tempPath);
    }

    [Fact]
    public void Validate_ValidTemplate_DoesNotThrow()
    {
        // Arrange
        CreateTestTemplate(new[] { "Sheet1", "Main" }, new Dictionary<string, string> { ["JobNumber"] = "Sheet1!$A$1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "A1", 0) },
            OutputFields = { new FieldDefinition("Main", "B2", 0) },
            HeaderInputs = { ["JobNumber"] = "123" }
        };

        // Act & Assert
        TemplateValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_MissingSheet_ThrowsValidationException()
    {
        // Arrange
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("MissingSheet", "A1", 0) }
        };

        // Act & Assert
        var ex = Assert.Throws<ValidationException>(() => TemplateValidator.Validate(_tempPath, config));
        Assert.Contains("Missing sheet 'MissingSheet'", ex.Message);
    }

    [Fact]
    public void Validate_MissingNamedRange_ThrowsValidationException()
    {
        // Arrange
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            HeaderInputs = { ["MissingRange"] = "val" }
        };

        // Act & Assert
        var ex = Assert.Throws<ValidationException>(() => TemplateValidator.Validate(_tempPath, config));
        Assert.Contains("Missing named range 'MissingRange'", ex.Message);
    }

    [Fact]
    public void Validate_FileInUse_CanStillRead()
    {
        // Arrange
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("Sheet1", "A1", 0) }
        };

        // Open the file exclusively to simulate it being open in Excel
        using var fs = new FileStream(_tempPath, FileMode.Open, FileAccess.Read, FileShare.None);

        // Act & Assert
        // This should NOT throw IOException if we use FileShare.ReadWrite in the validator
        TemplateValidator.Validate(_tempPath, config);
    }

    [Fact]
    public void Validate_AggregatesMultipleErrors()
    {
        // Arrange
        CreateTestTemplate(new[] { "Sheet1" });
        var config = new BatchConfig
        {
            InputFields = { new FieldDefinition("MissingSheet", "A1", 0) },
            HeaderInputs = { ["MissingRange"] = "val" }
        };

        // Act & Assert
        var ex = Assert.Throws<ValidationException>(() => TemplateValidator.Validate(_tempPath, config));
        Assert.Contains("Missing sheet 'MissingSheet'", ex.Message);
        Assert.Contains("Missing named range 'MissingRange'", ex.Message);
    }
}
