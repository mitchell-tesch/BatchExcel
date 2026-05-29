using BatchExcel.Models;

namespace BatchExcel.Tests;

public class BatchConfigTests
{
    [Fact]
    public void Macros_EmptyString_ReturnsEmptyList()
    {
        var config = new BatchConfig { CalculationMacrosRaw = "" };
        Assert.Empty(config.Macros);
    }

    [Fact]
    public void Macros_NullEquivalent_ReturnsEmptyList()
    {
        var config = new BatchConfig { CalculationMacrosRaw = "   " };
        Assert.Empty(config.Macros);
    }

    [Fact]
    public void Macros_SingleMacro_ReturnsOneItem()
    {
        var config = new BatchConfig { CalculationMacrosRaw = "MyMacro" };
        Assert.Equal(new[] { "MyMacro" }, config.Macros);
    }

    [Fact]
    public void Macros_CsvList_SplitsAndTrims()
    {
        var config = new BatchConfig { CalculationMacrosRaw = "FirstMacro, SecondMacro,  ThirdMacro" };
        Assert.Equal(new[] { "FirstMacro", "SecondMacro", "ThirdMacro" }, config.Macros);
    }

    [Fact]
    public void Macros_RemovesEmptyEntries()
    {
        var config = new BatchConfig { CalculationMacrosRaw = "A,,B, ,C" };
        Assert.Equal(new[] { "A", "B", "C" }, config.Macros);
    }

    [Fact]
    public void IncludedRunCount_CountsOnlyIncludedRuns()
    {
        var config = new BatchConfig
        {
            Calculations =
            {
                new BatchRun { Include = true },
                new BatchRun { Include = false },
                new BatchRun { Include = true },
                new BatchRun { Include = true },
            }
        };

        Assert.Equal(3, config.IncludedCalcCount);
    }

    [Fact]
    public void IncludedRunCount_EmptyRuns_ReturnsZero()
    {
        var config = new BatchConfig();
        Assert.Equal(0, config.IncludedCalcCount);
    }
}

