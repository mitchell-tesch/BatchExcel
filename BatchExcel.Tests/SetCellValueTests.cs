using DocumentFormat.OpenXml.Spreadsheet;
using BatchExcel.Services;

namespace BatchExcel.Tests;

/// <summary>
/// Tests OpenXmlHelpers.SetCellValue using in-memory SheetData (no file I/O required).
/// </summary>
public class SetCellValueTests
{
    [Fact]
    public void SetCellValue_CreatesRowAndCellWhenMissing()
    {
        var sheetData = new SheetData();

        OpenXmlHelpers.SetCellValue(sheetData, 5, 2, "hello");

        var row = sheetData.Elements<Row>().Single();
        Assert.Equal(5u, row.RowIndex?.Value);

        var cell = row.Elements<Cell>().Single();
        Assert.Equal("B5", cell.CellReference?.Value);
        Assert.Equal("hello", cell.CellValue?.Text);
        Assert.Equal(CellValues.String, cell.DataType?.Value);
    }

    [Fact]
    public void SetCellValue_NumericValue_NoDataTypeAttribute()
    {
        var sheetData = new SheetData();

        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, 42.5);

        var cell = sheetData.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("42.5", cell.CellValue?.Text);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void SetCellValue_BooleanValue_UsesBooleanType()
    {
        var sheetData = new SheetData();

        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, true);

        var cell = sheetData.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("1", cell.CellValue?.Text);
        Assert.Equal(CellValues.Boolean, cell.DataType?.Value);
    }

    [Fact]
    public void SetCellValue_NullValue_ClearsCellContents()
    {
        var sheetData = new SheetData();
        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, "initial");
        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, null);

        var cell = sheetData.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Null(cell.CellValue);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void SetCellValue_MultipleCellsInRow_PreservesColumnOrder()
    {
        var sheetData = new SheetData();

        // Insert out of column order
        OpenXmlHelpers.SetCellValue(sheetData, 1, 3, "C1");
        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, "A1");
        OpenXmlHelpers.SetCellValue(sheetData, 1, 2, "B1");

        var cells = sheetData.Elements<Row>().Single().Elements<Cell>().ToList();
        Assert.Equal(new[] { "A1", "B1", "C1" }, cells.Select(c => c.CellReference?.Value));
    }

    [Fact]
    public void SetCellValue_MultipleRows_PreservesRowOrder()
    {
        var sheetData = new SheetData();

        // Insert out of row order
        OpenXmlHelpers.SetCellValue(sheetData, 10, 1, "row10");
        OpenXmlHelpers.SetCellValue(sheetData, 5, 1, "row5");
        OpenXmlHelpers.SetCellValue(sheetData, 7, 1, "row7");

        var rows = sheetData.Elements<Row>().ToList();
        Assert.Equal(new uint[] { 5, 7, 10 }, rows.Select(r => r.RowIndex!.Value));
    }

    [Fact]
    public void SetCellValue_UpdatesExistingCell()
    {
        var sheetData = new SheetData();
        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, "first");
        OpenXmlHelpers.SetCellValue(sheetData, 1, 1, "second");

        var cell = sheetData.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("second", cell.CellValue?.Text);
    }

    [Fact]
    public void SetCellValue_ByRef_WorksIdenticallyToRowCol()
    {
        var byRef = new SheetData();
        var byRowCol = new SheetData();

        OpenXmlHelpers.SetCellValue(byRef, "C7", "value");
        OpenXmlHelpers.SetCellValue(byRowCol, 7, 3, "value");

        var refCell = byRef.Elements<Row>().Single().Elements<Cell>().Single();
        var rcCell = byRowCol.Elements<Row>().Single().Elements<Cell>().Single();

        Assert.Equal(refCell.CellReference?.Value, rcCell.CellReference?.Value);
        Assert.Equal(refCell.CellValue?.Text, rcCell.CellValue?.Text);
    }
}

