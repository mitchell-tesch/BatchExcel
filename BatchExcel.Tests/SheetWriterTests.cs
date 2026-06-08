using DocumentFormat.OpenXml.Spreadsheet;
using BatchExcel.Services;

namespace BatchExcel.Tests;

/// <summary>
/// Tests for the indexed SheetWriter — verifies cell ordering, row insertion,
/// updates of pre-existing cells, and parity with OpenXmlHelpers.SetCellValue.
/// </summary>
public class SheetWriterTests
{
    [Fact]
    public void SetCellValue_CreatesRowAndCell()
    {
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        writer.SetCellValue(5, 2, "hello");

        var row = sheet.Elements<Row>().Single();
        Assert.Equal(5u, row.RowIndex?.Value);
        var cell = row.Elements<Cell>().Single();
        Assert.Equal("B5", cell.CellReference?.Value);
        Assert.Equal("hello", cell.CellValue?.Text);
    }

    [Fact]
    public void SetCellValue_OutOfOrderRows_KeepsRowsSorted()
    {
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        writer.SetCellValue(10, 1, "row10");
        writer.SetCellValue(5, 1, "row5");
        writer.SetCellValue(7, 1, "row7");

        var rows = sheet.Elements<Row>().Select(r => r.RowIndex!.Value).ToArray();
        Assert.Equal(new uint[] { 5, 7, 10 }, rows);
    }

    [Fact]
    public void SetCellValue_OutOfOrderColumns_KeepsCellsSorted()
    {
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        writer.SetCellValue(1, 3, "C1");
        writer.SetCellValue(1, 1, "A1");
        writer.SetCellValue(1, 2, "B1");

        var refs = sheet.Elements<Row>().Single().Elements<Cell>()
            .Select(c => c.CellReference?.Value).ToArray();
        Assert.Equal(new[] { "A1", "B1", "C1" }, refs);
    }

    [Fact]
    public void SetCellValue_MultiLetterColumns_SortedByExcelOrderNotLexicographic()
    {
        // Excel orders columns by index: B (2) < AA (27) < AB (28). Lexicographic order would
        // put AA / AB before B. Catches regressions in SheetWriter's column-index comparator.
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        writer.SetCellValue(1, 28, "AB1");
        writer.SetCellValue(1, 27, "AA1");
        writer.SetCellValue(1, 2, "B1");
        writer.SetCellValue(1, 1, "A1");

        var refs = sheet.Elements<Row>().Single().Elements<Cell>()
            .Select(c => c.CellReference?.Value).ToArray();
        Assert.Equal(new[] { "A1", "B1", "AA1", "AB1" }, refs);
    }

    [Fact]
    public void SetCellValue_UpdatesExistingCell()
    {
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        writer.SetCellValue(1, 1, "first");
        writer.SetCellValue(1, 1, "second");

        var cell = sheet.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("second", cell.CellValue?.Text);
    }

    [Fact]
    public void SetCellValue_PicksUpExistingRowsAndCellsAtConstruction()
    {
        var sheet = new SheetData();
        // Pre-seed via the static helper
        OpenXmlHelpers.SetCellValue(sheet, 5, 2, "seeded");

        var writer = new SheetWriter(sheet);
        writer.SetCellValue(5, 2, "updated"); // should update existing cell, not append a new one

        var row = sheet.Elements<Row>().Single();
        var cell = row.Elements<Cell>().Single();
        Assert.Equal("updated", cell.CellValue?.Text);
    }

    [Fact]
    public void SetCellValue_ManyWrites_StaysFast()
    {
        // Performance regression guard: 10k writes should complete quickly with the index.
        var sheet = new SheetData();
        var writer = new SheetWriter(sheet);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 1; i <= 5000; i++)
        {
            writer.SetCellValue(i, 1, (double)i);
            writer.SetCellValue(i, 2, (double)i * 2);
        }
        sw.Stop();

        // Should comfortably complete in well under 1s on any modern machine.
        Assert.True(sw.ElapsedMilliseconds < 2000, $"SheetWriter too slow: {sw.ElapsedMilliseconds} ms");
        Assert.Equal(5000, sheet.Elements<Row>().Count());
    }
}

/// <summary>
/// Additional type coverage for OpenXmlHelpers.SetCellTypedValue via the public SetCellValue surface.
/// </summary>
public class SetCellTypedValueAdditionalTests
{
    [Fact]
    public void SetCellValue_Long_WritesAsNumber()
    {
        var sheet = new SheetData();
        OpenXmlHelpers.SetCellValue(sheet, 1, 1, 1234567890123L);

        var cell = sheet.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("1234567890123", cell.CellValue?.Text);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void SetCellValue_Decimal_WritesAsNumberInvariantCulture()
    {
        var sheet = new SheetData();
        OpenXmlHelpers.SetCellValue(sheet, 1, 1, 12.34m);

        var cell = sheet.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("12.34", cell.CellValue?.Text);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void SetCellValue_Float_WritesAsNumber()
    {
        var sheet = new SheetData();
        OpenXmlHelpers.SetCellValue(sheet, 1, 1, 0.5f);

        var cell = sheet.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal("0.5", cell.CellValue?.Text);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void SetCellValue_DateTime_WritesOADate()
    {
        var sheet = new SheetData();
        var dt = new DateTime(2024, 1, 1);
        OpenXmlHelpers.SetCellValue(sheet, 1, 1, dt);

        var cell = sheet.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal(dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
                     cell.CellValue?.Text);
    }
}

