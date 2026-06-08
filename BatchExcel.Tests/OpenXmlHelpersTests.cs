using BatchExcel.Services;

namespace BatchExcel.Tests;

public class OpenXmlHelpersTests
{
    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(1, 2, "B1")]
    [InlineData(5, 2, "B5")]
    [InlineData(1, 26, "Z1")]
    [InlineData(1, 27, "AA1")]
    [InlineData(1, 28, "AB1")]
    [InlineData(1, 52, "AZ1")]
    [InlineData(1, 53, "BA1")]
    [InlineData(100, 702, "ZZ100")] // 26 * 27
    [InlineData(1048576, 16384, "XFD1048576")] // Excel max cell
    public void GetCellReference_ProducesCorrectExcelReference(int row, int col, string expected)
    {
        Assert.Equal(expected, OpenXmlHelpers.GetCellReference(row, col));
    }

    [Theory]
    [InlineData("A1", 1u)]
    [InlineData("B5", 5u)]
    [InlineData("AA10", 10u)]
    [InlineData("XFD1048576", 1048576u)]
    public void ParseRowIndex_ExtractsRowNumber(string cellRef, uint expected)
    {
        Assert.Equal(expected, OpenXmlHelpers.ParseRowIndex(cellRef));
    }

    [Theory]
    [InlineData("A1", 1)]
    [InlineData("B5", 2)]
    [InlineData("Z1", 26)]
    [InlineData("AA1", 27)]
    [InlineData("AB1", 28)]
    [InlineData("AZ1", 52)]
    [InlineData("BA1", 53)]
    [InlineData("ZZ100", 702)]
    [InlineData("XFD1048576", 16384)] // Excel max column
    [InlineData("$B$5", 2)] // absolute refs: bails out at '$' which is fine for our use
    public void ParseColumnIndex_ExtractsColumnNumber(string cellRef, int expected)
    {
        Assert.Equal(expected, OpenXmlHelpers.ParseColumnIndex(cellRef));
    }

    [Fact]
    public void ParseColumnIndex_OrdersCellsByExcelColumnNotLexicographically()
    {
        // Critical: Excel orders "B" (col 2) before "AA" (col 27) but lexicographic ordering
        // puts "AA" first. SheetWriter relies on ParseColumnIndex to get this right.
        Assert.True(OpenXmlHelpers.ParseColumnIndex("B1") < OpenXmlHelpers.ParseColumnIndex("AA1"));
        Assert.True(string.Compare("AA1", "B1", System.StringComparison.Ordinal) < 0); // lexicographic disagrees
    }
}

