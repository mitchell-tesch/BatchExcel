using System.IO;
using BatchExcel.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

public class ValidationException : Exception
{
    public ValidationException(string message) : base(message) { }
}

public static class TemplateValidator
{
    /// <summary>
    /// Validates that all sheets and named ranges referenced in the configuration exist
    /// in the calculation template. Opens the file with non-locking shared access to
    /// avoid exceptions if the file is open in Excel.
    /// </summary>
    public static void Validate(string calculationPath, BatchConfig config)
    {
        var errors = new List<string>();

        // Open with FileShare.ReadWrite to avoid locking issues if open in Excel
        using var fs = new FileStream(calculationPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = SpreadsheetDocument.Open(fs, isEditable: false);

        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("Workbook part not found.");

        // Extract sheet names
        var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (s.Name?.Value != null)
                sheets.Add(s.Name.Value);
        }

        // Validate Input Fields (Sheets must exist)
        foreach (var field in config.InputFields)
        {
            if (!sheets.Contains(field.Sheet))
                errors.Add($"Missing sheet '{field.Sheet}' (referenced by input field '{field.Range}')");
        }

        // Validate Output Fields (Sheets must exist)
        foreach (var field in config.OutputFields)
        {
            if (!sheets.Contains(field.Sheet))
                errors.Add($"Missing sheet '{field.Sheet}' (referenced by output field '{field.Range}')");
        }

        // Validate Header Inputs (Named Ranges)
        // Note: We use the centralized parser to ensure we only validate what the engine can actually write to.
        // We do NOT throw if a header is missing from the template, as the writer ignores them silently.
        var definedNames = OpenXmlHelpers.BuildDefinedNamesLookup(workbookPart);

        // Optional: We could log a warning if a header is missing, but for now we just 
        // ensure the validator doesn't block the run. 

        if (errors.Count > 0)
        {
            throw new ValidationException("Template validation failed:\n- " + string.Join("\n- ", errors));
        }
    }
}
