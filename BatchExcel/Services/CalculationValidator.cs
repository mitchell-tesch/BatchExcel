using System.IO;
using System.Text.RegularExpressions;
using BatchExcel.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

/// <summary>
/// Thrown when the calculation workbook fails dry-run validation (missing sheets,
/// unresolved ranges/named ranges, corrupt file, etc.). Carries a user-friendly message
/// that the UI surfaces directly without a stack trace.
/// </summary>
public class ValidationException : Exception
{
    public ValidationException(string message) : base(message) { }
    public ValidationException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// Dry-run validator for the calculation workbook. Checks that every sheet and range
/// referenced by <see cref="BatchConfig.InputFields"/> / <see cref="BatchConfig.OutputFields"/>
/// can be resolved before the engine spins up Excel workers.
/// Header inputs (named ranges) are intentionally NOT enforced — the writer ignores
/// missing ones, matching <see cref="CalculationHeaderWriter"/> behavior.
/// </summary>
public static class CalculationValidator
{
    // Matches A1-style references and ranges: A1, $A$1, A1:B10, $A$1:$B$10, AA100:AB200, etc.
    // Case-insensitive; allows optional absolute-reference dollar signs on either side.
    private static readonly Regex A1RangePattern = new(
        @"^\$?[A-Za-z]{1,3}\$?[0-9]+(:\$?[A-Za-z]{1,3}\$?[0-9]+)?$",
        RegexOptions.Compiled);

    /// <summary>
    /// Validates the calculation template. Opens the file with shared read access so
    /// it works even while the user has the workbook open in Excel.
    /// Throws <see cref="ValidationException"/> with an aggregated error list on failure.
    /// </summary>
    public static void Validate(string calculationPath, BatchConfig config)
    {
        FileStream fs;
        try
        {
            // FileShare.ReadWrite so we don't conflict with Excel if the user has it open.
            fs = new FileStream(calculationPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch (IOException ex)
        {
            throw new ValidationException(
                $"Could not open calculation template '{calculationPath}': {ex.Message}", ex);
        }

        try
        {
            SpreadsheetDocument document;
            try
            {
                document = SpreadsheetDocument.Open(fs, isEditable: false);
            }
            catch (Exception ex) when (ex is not ValidationException)
            {
                throw new ValidationException(
                    $"Calculation template '{calculationPath}' is not a valid .xlsx file: {ex.Message}", ex);
            }

            using (document)
            {
                var workbookPart = document.WorkbookPart
                                   ?? throw new ValidationException(
                                       $"Calculation template '{calculationPath}' has no workbook part.");

                var sheets = BuildSheetNameSet(workbookPart);
                var definedNames = OpenXmlHelpers.BuildDefinedNamesLookup(workbookPart);

                var errors = new List<string>();
                ValidateFields(config.InputFields, "input", sheets, definedNames, errors);
                ValidateFields(config.OutputFields, "output", sheets, definedNames, errors);

                if (errors.Count > 0)
                {
                    throw new ValidationException(
                        "Calculation template validation failed:\n  - " + string.Join("\n  - ", errors));
                }
            }
        }
        finally
        {
            fs.Dispose();
        }
    }

    private static HashSet<string> BuildSheetNameSet(WorkbookPart workbookPart)
    {
        var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (s.Name?.Value is { } name)
                sheets.Add(name);
        }
        return sheets;
    }

    private static void ValidateFields(
        List<FieldDefinition> fields,
        string fieldKind,
        HashSet<string> sheets,
        Dictionary<string, (string sheetName, string cellRef)> definedNames,
        List<string> errors)
    {
        foreach (var field in fields)
        {
            // Sheet must exist
            if (!sheets.Contains(field.Sheet))
            {
                errors.Add($"Missing sheet '{field.Sheet}' (referenced by {fieldKind} field range '{field.Range}')");
                continue; // skip range check — without the sheet, range info is moot
            }

            // Range must be either an A1-style reference or a known defined name.
            if (string.IsNullOrWhiteSpace(field.Range))
            {
                errors.Add($"Empty range on sheet '{field.Sheet}' for {fieldKind} field");
                continue;
            }

            if (A1RangePattern.IsMatch(field.Range))
                continue;

            if (definedNames.ContainsKey(field.Range))
                continue;

            errors.Add(
                $"Range '{field.Range}' on sheet '{field.Sheet}' for {fieldKind} field is not a valid A1 " +
                "reference and is not a defined name in the calculation template");
        }
    }
}

