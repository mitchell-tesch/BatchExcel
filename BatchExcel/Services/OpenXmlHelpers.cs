using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BatchExcel.Services;

/// <summary>
/// Shared helpers for OpenXML SpreadsheetML cell manipulation.
/// </summary>
internal static class OpenXmlHelpers
{
    /// <summary>
    /// Builds a lookup of defined name → (sheet name, cell reference).
    /// Parses references like "SheetName!$B$5" into their components.
    /// Returns ONLY valid, writable cell references; offsets and formulas are ignored.
    /// </summary>
    public static Dictionary<string, (string sheetName, string cellRef)> BuildDefinedNamesLookup(WorkbookPart workbookPart)
    {
        var definedNames = new Dictionary<string, (string sheetName, string cellRef)>(StringComparer.OrdinalIgnoreCase);

        if (workbookPart.Workbook.DefinedNames == null)
            return definedNames;

        foreach (var dn in workbookPart.Workbook.DefinedNames.Elements<DefinedName>())
        {
            string? name = dn.Name;
            var reference = dn.Text;
            if (name == null) continue;

            // Strict parsing: expects "Sheet!$A$1"
            var parts = reference.Split('!');
            if (parts.Length != 2) continue;
            var sName = parts[0].Trim('\'');
            var cRef = parts[1].Replace("$", "");
            definedNames[name] = (sName, cRef);
        }

        return definedNames;
    }

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
    /// Parses the column index from a cell reference string (e.g., "B5" → 2, "AA1" → 27).
    /// Tolerates a leading absolute-reference '$' (e.g. "$B$5" → 2). Used to sort cells within
    /// a row by Excel column order (which differs from lexicographic order: "B" &lt; "AA" by
    /// column number but "AA" &lt; "B" lexicographically).
    /// </summary>
    public static int ParseColumnIndex(string cellRef)
    {
        int col = 0;
        for (int i = 0; i < cellRef.Length; i++)
        {
            var ch = cellRef[i];
            if (ch == '$' && col == 0) continue; // skip leading absolute marker
            if (ch >= 'A' && ch <= 'Z')
                col = col * 26 + (ch - 'A' + 1);
            else if (ch >= 'a' && ch <= 'z')
                col = col * 26 + (ch - 'a' + 1);
            else
                break; // hit '$' before row digits, or the row digits themselves
        }
        return col;
    }

    /// <summary>
    /// Sets a cell value at a specific row/col position, creating the row/cell if missing.
    /// Preserves cell ordering within the row.
    /// <para>
    /// <b>Performance:</b> O(rows + cells-in-row) per call due to the LINQ <c>FirstOrDefault</c>
    /// scans. Suitable for one-off writes (and as a test helper). For bulk writes, use
    /// <see cref="SheetWriter"/> which indexes both dimensions for O(1) lookup.
    /// </para>
    /// </summary>
    public static void SetCellValue(SheetData sheetData, int rowIndex, int colIndex, object? value)
    {
        string cellRef = GetCellReference(rowIndex, colIndex);
        SetCellValue(sheetData, cellRef, (uint)rowIndex, value);
    }

    /// <summary>
    /// Sets a cell value by cell reference string (e.g., "B5"), creating the row/cell if missing.
    /// See the row/col overload for performance notes — prefer <see cref="SheetWriter"/> for
    /// bulk writes.
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
/// Builds row and per-row cell indices on first use so cell lookups are O(1) instead of
/// O(rows × cells). Cells within a row are sorted by Excel column index (not lexicographic
/// order — Excel orders "B" before "AA" by column number, lexicographic sort would put "AA"
/// first). Append-at-end inserts (the common case: writing rows top-to-bottom, columns
/// left-to-right) are detected via O(1) max-tracking and short-circuit to a direct Append
/// without scanning for an insertion point. Not thread-safe — use one instance per worksheet
/// per thread.
/// </summary>
internal sealed class SheetWriter
{
    private readonly SheetData _sheetData;
    private readonly SortedDictionary<uint, Row> _rows = new();

    // Per-row state: column-index → Cell (sorted) plus a cached max column index for O(1)
    // append-fast-path detection. SortedDictionary doesn't expose Max in O(log n), and walking
    // .Keys is O(n), so we maintain _maxCol explicitly.
    private sealed class RowCells
    {
        public readonly SortedDictionary<int, Cell> Map = new();
        public int MaxCol = -1; // -1 = empty
    }

    private readonly Dictionary<uint, RowCells> _cellsByRow = new();

    // O(1)-tracked max row index across the whole sheet.
    private uint _maxRow;
    private bool _hasAnyRow;

    public SheetWriter(SheetData sheetData)
    {
        _sheetData = sheetData;

        // Index existing rows and cells in a single pass
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex?.Value is not { } idx) continue;
            _rows[idx] = row;
            if (!_hasAnyRow || idx > _maxRow) { _maxRow = idx; _hasAnyRow = true; }

            var rowCells = new RowCells();
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } cr)
                {
                    var col = OpenXmlHelpers.ParseColumnIndex(cr);
                    if (col > 0)
                    {
                        rowCells.Map[col] = cell;
                        if (col > rowCells.MaxCol) rowCells.MaxCol = col;
                    }
                }
            }
            _cellsByRow[idx] = rowCells;
        }
    }

    public void SetCellValue(int rowIndex, int colIndex, object? value)
    {
        string cellRef = OpenXmlHelpers.GetCellReference(rowIndex, colIndex);
        SetCellValue(cellRef, (uint)rowIndex, colIndex, value);
    }

    public void SetCellValue(string cellRef, object? value)
    {
        uint rowIndex = OpenXmlHelpers.ParseRowIndex(cellRef);
        int colIndex = OpenXmlHelpers.ParseColumnIndex(cellRef);
        SetCellValue(cellRef, rowIndex, colIndex, value);
    }

    private void SetCellValue(string cellRef, uint rowIndex, int colIndex, object? value)
    {
        if (!_rows.TryGetValue(rowIndex, out var row))
        {
            row = new Row { RowIndex = rowIndex };

            // Fast path: appending past the current max row (the dominant workload — we usually
            // write rows top-to-bottom). O(1) thanks to _maxRow tracking.
            if (!_hasAnyRow || rowIndex > _maxRow)
            {
                _sheetData.Append(row);
            }
            else
            {
                // Slow path: out-of-order row insert. Walk the sorted dictionary to find the
                // first row with a greater index. Cost is O(insertion-point distance); rare in
                // practice because writes flow top-to-bottom.
                Row? refRow = null;
                foreach (var kvp in _rows)
                {
                    if (kvp.Key > rowIndex) { refRow = kvp.Value; break; }
                }
                if (refRow != null) _sheetData.InsertBefore(row, refRow);
                else _sheetData.Append(row);
            }

            _rows[rowIndex] = row;
            if (!_hasAnyRow || rowIndex > _maxRow) { _maxRow = rowIndex; _hasAnyRow = true; }
            _cellsByRow[rowIndex] = new RowCells();
        }

        var rowCells = _cellsByRow[rowIndex];
        if (!rowCells.Map.TryGetValue(colIndex, out var cell))
        {
            cell = new Cell { CellReference = cellRef };

            // Same O(1) fast path for columns within a row.
            if (rowCells.MaxCol < 0 || colIndex > rowCells.MaxCol)
            {
                row.Append(cell);
            }
            else
            {
                Cell? refCell = null;
                foreach (var kvp in rowCells.Map)
                {
                    if (kvp.Key > colIndex) { refCell = kvp.Value; break; }
                }
                if (refCell != null) row.InsertBefore(cell, refCell);
                else row.Append(cell);
            }

            rowCells.Map[colIndex] = cell;
            if (colIndex > rowCells.MaxCol) rowCells.MaxCol = colIndex;
        }

        OpenXmlHelpers.SetCellTypedValue(cell, value);
    }
}
