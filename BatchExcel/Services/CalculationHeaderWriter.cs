using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

/// <summary>
/// Writes header information (Date, JobNumber, Project, etc.) to the calculation workbook
/// using the OpenXML SDK directly — no Excel process required.
/// Resolves Excel defined names to their cell references for accurate placement.
/// </summary>
public static class CalculationHeaderWriter
{
    private const string DateRangeName = "Date";

    /// <summary>
    /// Writes header values into the calculation file. Header values are placed at named ranges
    /// matching the dictionary keys (e.g., "JobNumber" → range named "JobNumber" in the calculation).
    /// The current batch start time is written to a range named "Date".
    /// </summary>
    public static void Write(
        string calculationPath,
        Dictionary<string, object?> headerInputs,
        DateTime batchStart)
    {

        // Retry the open to tolerate transient sharing violations on SMB shares — the staged
        // calc file was just written by File.Copy and AV/Indexer may briefly hold it.
        using var document = IoRetry.Run(() => SpreadsheetDocument.Open(calculationPath, isEditable: true));
        var workbookPart = document.WorkbookPart
                           ?? throw new InvalidOperationException("Workbook part not found.");

        var definedNames = BuildDefinedNamesLookup(workbookPart);
        var sheetLookup = BuildWorksheetLookup(workbookPart);

        // Cache one SheetWriter per worksheet so multiple writes to the same sheet share an index.
        var writers = new Dictionary<WorksheetPart, SheetWriter>();

        // Write "Date" named range
        if (definedNames.TryGetValue(DateRangeName, out var dateInfo) &&
            sheetLookup.TryGetValue(dateInfo.sheetName, out var datePart))
        {
            GetWriter(datePart).SetCellValue(dateInfo.cellRef, batchStart.ToString("G"));
        }

        // Write each header field to its corresponding named range
        foreach (var (rangeName, value) in headerInputs)
        {
            if (definedNames.TryGetValue(rangeName, out var info) &&
                sheetLookup.TryGetValue(info.sheetName, out var wsPart))
            {
                GetWriter(wsPart).SetCellValue(info.cellRef, value);
            }
        }

        // Save only the worksheets we actually modified
        foreach (var wsPart in writers.Keys)
        {
            wsPart.Worksheet.Save();
        }

        return;

        SheetWriter GetWriter(WorksheetPart wsPart)
        {
            if (writers.TryGetValue(wsPart, out var w)) return w;
            w = new SheetWriter(wsPart.Worksheet.GetFirstChild<SheetData>()!);
            writers[wsPart] = w;
            return w;
        }
    }

    /// <summary>
    /// Builds a lookup of defined name → (sheet name, cell reference).
    /// Parses references like "SheetName!$B$5" into their components.
    /// </summary>
    private static Dictionary<string, (string sheetName, string cellRef)> BuildDefinedNamesLookup(WorkbookPart workbookPart)
    {
        var definedNames = new Dictionary<string, (string sheetName, string cellRef)>(StringComparer.OrdinalIgnoreCase);

        if (workbookPart.Workbook.DefinedNames == null)
            return definedNames;

        foreach (var dn in workbookPart.Workbook.DefinedNames.Elements<DefinedName>())
        {
            string? name = dn.Name;
            var reference = dn.Text;
            if (name == null) continue;

            var parts = reference.Split('!');
            if (parts.Length != 2) continue;
            var sName = parts[0].Trim('\'');
            var cRef = parts[1].Replace("$", "");
            definedNames[name] = (sName, cRef);
        }

        return definedNames;
    }

    /// <summary>
    /// Builds a lookup of sheet name → WorksheetPart, skipping chart sheets and other non-worksheet parts.
    /// </summary>
    private static Dictionary<string, WorksheetPart> BuildWorksheetLookup(WorkbookPart workbookPart)
    {
        var lookup = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);

        foreach (var sheetEntry in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (sheetEntry.Name?.Value != null && sheetEntry.Id?.Value != null)
            {
                var part = workbookPart.GetPartById(sheetEntry.Id.Value);
                if (part is WorksheetPart wsPart)
                    lookup[sheetEntry.Name.Value] = wsPart;
            }
        }

        return lookup;
    }
}

