using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

/// <summary>
/// Shared helpers for OpenXML SpreadsheetML cell manipulation.
/// </summary>
internal static class OpenXmlHelpers
{
    /// <summary>
    /// Converts 1-based row and column indices to an Excel cell reference (e.g., row=1, col=2 → "B1").
    /// </summary>
    public static string GetCellReference(int row, int col)
    {
        string colLetter = "";
        int c = col;
        while (c > 0)
        {
            c--;
            colLetter = (char)('A' + c % 26) + colLetter;
            c /= 26;
        }
        return colLetter + row;
    }

    /// <summary>
    /// Parses the row number from a cell reference string (e.g., "B5" → 5).
    /// </summary>
    public static uint ParseRowIndex(string cellRef)
    {
        int rowStart = 0;
        for (int i = 0; i < cellRef.Length; i++)
        {
            if (char.IsDigit(cellRef[i]))
            {
                rowStart = i;
                break;
            }
        }
        return uint.Parse(cellRef[rowStart..]);
    }

    /// <summary>
    /// Sets a cell value at a specific row/col position, creating the row/cell if missing.
    /// Preserves cell ordering within the row.
    /// </summary>
    public static void SetCellValue(SheetData sheetData, int rowIndex, int colIndex, object? value)
    {
        string cellRef = GetCellReference(rowIndex, colIndex);
        SetCellValue(sheetData, cellRef, (uint)rowIndex, value);
    }

    /// <summary>
    /// Sets a cell value by cell reference string (e.g., "B5"), creating the row/cell if missing.
    /// </summary>
    public static void SetCellValue(SheetData sheetData, string cellRef, object? value)
    {
        uint rowIndex = ParseRowIndex(cellRef);
        SetCellValue(sheetData, cellRef, rowIndex, value);
    }

    private static void SetCellValue(SheetData sheetData, string cellRef, uint rowIndex, object? value)
    {
        // Find or create the row (insert in row-index order)
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            var refRow = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value > rowIndex);
            if (refRow != null)
                sheetData.InsertBefore(row, refRow);
            else
                sheetData.Append(row);
        }

        // Find or create the cell (insert in column order)
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef };
            var refCell = row.Elements<Cell>()
                .FirstOrDefault(c => string.Compare(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase) > 0);
            if (refCell != null)
                row.InsertBefore(cell, refCell);
            else
                row.Append(cell);
        }

        // Set the typed value
        SetCellTypedValue(cell, value);
    }

    internal static void SetCellTypedValue(Cell cell, object? value)
    {
        if (value == null)
        {
            cell.CellValue = null;
            cell.DataType = null;
        }
        else if (value is double d)
        {
            cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
            cell.DataType = null; // numeric = no DataType attribute
        }
        else if (value is int i)
        {
            cell.CellValue = new CellValue(i.ToString(CultureInfo.InvariantCulture));
            cell.DataType = null;
        }
        else if (value is long l)
        {
            cell.CellValue = new CellValue(l.ToString(CultureInfo.InvariantCulture));
            cell.DataType = null;
        }
        else if (value is float f)
        {
            cell.CellValue = new CellValue(((double)f).ToString(CultureInfo.InvariantCulture));
            cell.DataType = null;
        }
        else if (value is decimal m)
        {
            cell.CellValue = new CellValue(m.ToString(CultureInfo.InvariantCulture));
            cell.DataType = null;
        }
        else if (value is bool b)
        {
            cell.CellValue = new CellValue(b ? "1" : "0");
            cell.DataType = CellValues.Boolean;
        }
        else if (value is DateTime dt)
        {
            cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
            cell.DataType = null;
        }
        else
        {
            cell.CellValue = new CellValue(value.ToString() ?? "");
            cell.DataType = CellValues.String;
        }
    }
}

/// <summary>
/// Indexed wrapper around a <see cref="SheetData"/> for fast bulk cell writes.
/// Builds row and per-row cell dictionaries on first use so cell writes are O(1) instead of O(rows×cells).
/// Not thread-safe — use one instance per worksheet per thread.
/// </summary>
internal sealed class SheetWriter
{
    private readonly SheetData _sheetData;
    private readonly SortedDictionary<uint, Row> _rows = new();
    private readonly Dictionary<uint, Dictionary<string, Cell>> _cells = new();

    public SheetWriter(SheetData sheetData)
    {
        _sheetData = sheetData;

        // Index existing rows and cells in a single pass
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex?.Value is not { } idx) continue;
            _rows[idx] = row;
            var cellMap = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } cr)
                    cellMap[cr] = cell;
            }
            _cells[idx] = cellMap;
        }
    }

    public void SetCellValue(int rowIndex, int colIndex, object? value)
    {
        string cellRef = OpenXmlHelpers.GetCellReference(rowIndex, colIndex);
        SetCellValue(cellRef, (uint)rowIndex, value);
    }

    public void SetCellValue(string cellRef, object? value)
    {
        uint rowIndex = OpenXmlHelpers.ParseRowIndex(cellRef);
        SetCellValue(cellRef, rowIndex, value);
    }

    private void SetCellValue(string cellRef, uint rowIndex, object? value)
    {
        if (!_rows.TryGetValue(rowIndex, out var row))
        {
            row = new Row { RowIndex = rowIndex };
            // Find insertion point using sorted key set (O(log n))
            Row? refRow = null;
            foreach (var kvp in _rows)
            {
                if (kvp.Key > rowIndex) { refRow = kvp.Value; break; }
            }
            if (refRow != null)
                _sheetData.InsertBefore(row, refRow);
            else
                _sheetData.Append(row);

            _rows[rowIndex] = row;
            _cells[rowIndex] = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        }

        var cellMap = _cells[rowIndex];
        if (!cellMap.TryGetValue(cellRef, out var cell))
        {
            cell = new Cell { CellReference = cellRef };
            // Find insertion point within row (still O(n) within row, but rows are typically narrow)
            Cell? refCell = null;
            foreach (var existing in row.Elements<Cell>())
            {
                if (string.Compare(existing.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    refCell = existing;
                    break;
                }
            }
            if (refCell != null)
                row.InsertBefore(cell, refCell);
            else
                row.Append(cell);

            cellMap[cellRef] = cell;
        }

        OpenXmlHelpers.SetCellTypedValue(cell, value);
    }
}
