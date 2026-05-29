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
}

