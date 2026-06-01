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

        // Extract defined names (named ranges)
        var definedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (workbookPart.Workbook.DefinedNames != null)
        {
            foreach (var dn in workbookPart.Workbook.DefinedNames.Elements<DefinedName>())
            {
                if (dn.Name?.Value != null)
                    definedNames.Add(dn.Name.Value);
            }
        }

        // Validate Input Fields
        foreach (var field in config.InputFields)
        {
            if (!sheets.Contains(field.Sheet))
                errors.Add($"Missing sheet '{field.Sheet}' (referenced by input field '{field.Range}')");
        }

        // Validate Output Fields
        foreach (var field in config.OutputFields)
        {
            if (!sheets.Contains(field.Sheet))
                errors.Add($"Missing sheet '{field.Sheet}' (referenced by output field '{field.Range}')");
        }

        // Validate Header Inputs (Named Ranges)
        foreach (var rangeName in config.HeaderInputs.Keys)
        {
            // "Date" is a special case written automatically, but it still needs a defined name
            if (!definedNames.Contains(rangeName))
                errors.Add($"Missing named range '{rangeName}' (referenced by project header)");
        }

        if (errors.Count > 0)
        {
            throw new ValidationException("Template validation failed:\n- " + string.Join("\n- ", errors));
        }
    }
}
