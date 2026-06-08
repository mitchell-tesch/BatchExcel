using System.IO;
using BatchExcel.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

/// <summary>
/// Reads batch configuration from and writes results to a BatchExcel batcher spreadsheet.
/// Uses ClosedXML for reading and raw OpenXML SDK for writing (to avoid ClosedXML's
/// theme color serialization bugs with conditional formatting in some spreadsheets).
/// No Excel interop is required.
/// </summary>
public static class BatcherReader
{
    // Batcher spreadsheet layout constants
    private const string BatchInSheet = "Main";
    private const int DataStartRow = 15;
    private const int DataStartCol = 2; // Column B

    /// <summary>
    /// Number of header rows in the data table: row 0=types, row 1=sheets, row 2=ranges. Run data starts at row 3.
    /// </summary>
    private const int DataHeaderRowCount = 3;

    private static readonly Dictionary<string, (int Row, int Col)> HeaderInRange = new()
    {
        ["calculation_file"] = (3, 2),
        ["calculation_macros"] = (4, 2),
        ["calculation_header-sheet"] = (5, 2),
        ["JobNumber"] = (7, 2),
        ["Project"] = (8, 2),
        ["Designer"] = (9, 2),
        ["DesignNotes"] = (10, 2),
    };

    /// <summary>
    /// Reads all batch configuration from the batcher workbook using ClosedXML (no Excel Interop needed).
    /// </summary>
    public static BatchConfig ReadConfig(string batcherFilePath)
    {
        var config = new BatchConfig();

        // Retry the open to tolerate transient sharing violations on SMB shares (AV scanners,
        // Search Indexer, OneDrive, etc.) — same failure mode as WriteResults.
        using var workbook = IoRetry.Run(() => new XLWorkbook(Path.GetFullPath(batcherFilePath)));
        var sheet = workbook.Worksheet(BatchInSheet);

        // Read calculation and header inputs
        var headerInputs = new Dictionary<string, object?>();
        foreach (var (key, pos) in HeaderInRange)
        {
            var cell = sheet.Cell(pos.Row, pos.Col);
            var strValue = cell.GetString();

            var prefix = key.Split('_')[0].ToLower();
            if (prefix == "calculation")
            {
                switch (key)
                {
                    case "calculation_file": config.CalculationFile = strValue; break;
                    case "calculation_macros": config.CalculationMacrosRaw = strValue; break;
                    case "calculation_header-sheet": config.CalculationHeaderSheet = strValue; break;
                }
            }
            else
            {
                headerInputs[key] = cell.Value.IsBlank ? null : GetCellValue(cell);
            }
        }
        config.HeaderInputs = headerInputs;

        // Find the extent of the data table starting at the configured start cell.
        // Use ClosedXML's row/column "last used cell" lookups (one walk over the used range)
        // instead of probing one cell at a time. Assumes the data table is contiguous from the
        // start cell — which has always been the layout contract for the batcher template.
        var headerRow = sheet.Row(DataStartRow);
        var lastColCell = headerRow.LastCellUsed();
        var lastCol = lastColCell != null
            ? Math.Max(DataStartCol, lastColCell.Address.ColumnNumber)
            : DataStartCol;

        var firstCol = sheet.Column(DataStartCol);
        var lastRowCell = firstCol.LastCellUsed();
        var lastRow = lastRowCell != null
            ? Math.Max(DataStartRow, lastRowCell.Address.RowNumber)
            : DataStartRow;

        var dataRows = lastRow - DataStartRow + 1;
        var dataCols = lastCol - DataStartCol + 1;

        if (dataRows < DataHeaderRowCount + 1 || dataCols < 2)
            throw new InvalidOperationException(
                $"Batch input table is too small. Expected at least {DataHeaderRowCount + 1} rows " +
                "(types, sheets, ranges, data) and 2 columns.");

        // Parse field definitions from the header rows (skip column 0 which is Include/Status)
        for (var c = 1; c < dataCols; c++)
        {
            var col = DataStartCol + c;
            var type = sheet.Cell(DataStartRow, col).GetString().Trim().ToLower();
            var fieldSheet = sheet.Cell(DataStartRow + 1, col).GetString().Trim();
            var fieldRange = sheet.Cell(DataStartRow + 2, col).GetString().Trim();

            var field = new FieldDefinition(fieldSheet, fieldRange, c);

            switch (type)
            {
                case "in": config.InputFields.Add(field); break;
                case "out": config.OutputFields.Add(field); break;
                default: config.SkipFields.Add(field); break;
            }
        }

        // Parse run data (rows after the header rows)
        for (var r = DataHeaderRowCount; r < dataRows; r++)
        {
            var row = DataStartRow + r;
            var status = sheet.Cell(row, DataStartCol).GetString().Trim();
            var include = status.Equals("Yes", StringComparison.OrdinalIgnoreCase);

            var titleStr = sheet.Cell(row, DataStartCol + 1).GetString().Trim();
            var title = string.IsNullOrEmpty(titleStr) ? $"Run {r - DataHeaderRowCount + 1}" : titleStr;

            var data = new object?[dataCols];
            for (var c = 0; c < dataCols; c++)
            {
                var cell = sheet.Cell(row, DataStartCol + c);
                data[c] = cell.Value.IsBlank ? null : GetCellValue(cell);
            }

            config.Calculations.Add(new BatchRun
            {
                Index = r - DataHeaderRowCount,
                Include = include,
                Title = title,
                Data = data
            });
        }

        return config;
    }

    /// <summary>
    /// Writes batch results back into the batcher workbook and saves a copy to the output folder.
    /// Writes the OpenXML modifications once to the copy, then mirrors it to the original.
    /// If the original cannot be overwritten (e.g., user has it open in Excel), the copy is
    /// preserved and an <see cref="IOException"/> is rethrown so the caller can warn.
    /// </summary>
    /// <returns>true if the original was successfully updated; false if only the copy was written.</returns>
    public static bool WriteResults(string batcherFilePath, BatchConfig config, string outputFolder)
    {
        var fullPath = Path.GetFullPath(batcherFilePath);
        var copyPath = Path.Combine(outputFolder, Path.GetFileName(batcherFilePath));

        // Copy original → output folder, modify the copy once, then mirror the modified copy
        // back over the original. Avoids running the OpenXML write twice.
        //
        // Each step is wrapped in IoRetry to tolerate transient SMB/AV sharing violations
        // that commonly hit a freshly-written .xlsx on a network share (antivirus scan, Search
        // Indexer, OneDrive, etc. briefly opening the file via SMB change notifications).
        IoRetry.Run(() => File.Copy(fullPath, copyPath, overwrite: true));
        IoRetry.Run(() => WriteResultsOpenXml(copyPath, config));

        try
        {
            IoRetry.Run(() => File.Copy(copyPath, fullPath, overwrite: true));
            return true;
        }
        catch (IOException)
        {
            // Original is locked (e.g. open in Excel) or still contested after all retries.
            // The copy in the output folder is intact.
            return false;
        }
    }

    private static void WriteResultsOpenXml(string filePath, BatchConfig config)
    {
        using var document = SpreadsheetDocument.Open(filePath, isEditable: true);
        var workbookPart = document.WorkbookPart
                           ?? throw new InvalidOperationException("Workbook part not found.");

        var sheetEntry = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == BatchInSheet)
            ?? throw new InvalidOperationException($"Sheet '{BatchInSheet}' not found.");

        var worksheetPart = workbookPart.GetPartById(sheetEntry.Id!) as WorksheetPart
                            ?? throw new InvalidOperationException($"Sheet '{BatchInSheet}' is not a worksheet.");
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        // Use indexed SheetWriter so cell lookups are O(1) instead of O(rows × cells).
        var writer = new SheetWriter(sheetData);

        for (var f = 0; f < config.OutputFields.Count; f++)
        {
            var field = config.OutputFields[f];
            var col = DataStartCol + field.ColumnOffset;

            foreach (var run in config.Calculations)
            {
                if (!run.Include || run.Results == null) continue;

                var row = DataStartRow + DataHeaderRowCount + run.Index;
                var value = run.Results.Length > f ? run.Results[f] : null;

                writer.SetCellValue(row, col, value);
            }
        }

        worksheetPart.Worksheet.Save();
    }

    private static object? GetCellValue(IXLCell cell)
    {
        if (cell.Value.IsBlank) return null;
        if (cell.Value.IsNumber) return cell.Value.GetNumber();
        if (cell.Value.IsBoolean) return cell.Value.GetBoolean();
        if (cell.Value.IsDateTime) return cell.Value.GetDateTime();
        return cell.Value.IsText ? cell.Value.GetText() : cell.GetString();
    }
}
