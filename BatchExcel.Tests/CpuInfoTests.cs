using BatchExcel.Services;

namespace BatchExcel.Tests;

public class CpuInfoTests
{
    [Fact]
    public void PhysicalCoreCount_IsWithinLogicalProcessorCount()
    {
        var cores = CpuInfo.PhysicalCoreCount();

        Assert.InRange(cores, 1, Environment.ProcessorCount);
    }
}
