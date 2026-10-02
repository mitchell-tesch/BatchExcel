using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using BatchExcel.Models;
using BatchExcel.Services;

namespace BatchExcel.Tests;

public class ExcelWorkerTests
{
    [Theory]
    [InlineData(unchecked((int)0x800706BA))] // RPC_S_SERVER_UNAVAILABLE
    [InlineData(unchecked((int)0x800706BE))] // RPC_S_CALL_FAILED
    [InlineData(unchecked((int)0x800706BF))] // RPC_S_CALL_FAILED_DNE
    [InlineData(unchecked((int)0x80010007))] // RPC_E_SERVER_DIED
    [InlineData(unchecked((int)0x80010012))] // RPC_E_SERVER_DIED_DNE
    [InlineData(unchecked((int)0x80010108))] // RPC_E_DISCONNECTED
    public void IsExcelUnavailable_DeadServerHResults_ReturnsTrue(int hr)
    {
        Assert.True(ExcelWorker.IsExcelUnavailable(new COMException("dead", hr)));
    }

    [Theory]
    [InlineData(unchecked((int)0x80010001))] // RPC_E_CALL_REJECTED (transient, Excel alive)
    [InlineData(unchecked((int)0x800AC472))] // VBA_E_IGNORE (Excel busy)
    [InlineData(unchecked((int)0x800A03EC))] // Generic Excel error (e.g. bad range / macro error)
    public void IsExcelUnavailable_OtherHResults_ReturnsFalse(int hr)
    {
        Assert.False(ExcelWorker.IsExcelUnavailable(new COMException("alive", hr)));
    }

    [Fact]
    public void RestoreBatchAppState_ResetsStateLeftByMacro()
    {
        var app = new FakeExcelApp { ScreenUpdating = true, EnableEvents = true, DisplayAlerts = true, Calculation = -4105 };

        ExcelWorker.RestoreBatchAppState(app);

        Assert.False(app.ScreenUpdating);
        Assert.False(app.EnableEvents);
        Assert.False(app.DisplayAlerts);
        Assert.Equal(-4135, app.Calculation);
    }

    public class FakeExcelApp
    {
        public bool ScreenUpdating { get; set; }
        public bool EnableEvents { get; set; }
        public bool DisplayAlerts { get; set; }
        public int Calculation { get; set; }
        public List<string> Calls { get; } = [];
        public void Calculate() => Calls.Add("Calculate");
        public object? Run(string macro) { Calls.Add($"Run:{macro}"); return null; }
    }

    public class FakeRange
    {
        public object? Value { get; set; }
    }

    [Theory]
    [InlineData(new[] { "Solve" }, new[] { "Calculate", "Run:Solve", "Calculate" })]
    [InlineData(new string[0], new[] { "Calculate" })]
    public void ProcessSingleRun_RecalculatesAfterMacros(string[] macros, string[] expectedCalls)
    {
        var run = new BatchRun { Index = 0, Include = true, Title = "r" };
        var config = new BatchConfig();
        config.Calculations.Add(run);
        var ctx = new WorkerContext(1, "", "", config, [.. macros], new ConcurrentQueue<BatchRun>(), "",
            SaveRuns: false, [], CancellationToken.None, _ => { }, () => { });
        var app = new FakeExcelApp();

        new ExcelWorker(ctx).ProcessSingleRun(app, null!, run, [], [new FakeRange { Value = 1.5 }]);

        Assert.Equal(expectedCalls, app.Calls);
        Assert.Equal([1.5], run.Results);
    }

    [Theory]
    [InlineData(-2146826281, "#DIV/0!")]
    [InlineData(-2146826246, "#N/A")]
    [InlineData(-2146826273, "#VALUE!")]
    [InlineData(-2146826265, "#REF!")]
    [InlineData(-2146826259, "#NAME?")]
    [InlineData(-2146826252, "#NUM!")]
    [InlineData(-2146826288, "#NULL!")]
    public void ExcelError_FromComValue_MapsErrorCodes(int comValue, string expected)
    {
        var result = Assert.IsType<BatchExcel.Models.ExcelError>(BatchExcel.Models.ExcelError.FromComValue(comValue));
        Assert.Equal(expected, result.Text);
    }

    [Theory]
    [InlineData(1.5)]
    [InlineData("text")]
    [InlineData(true)]
    [InlineData(42)]
    [InlineData(null)]
    public void ExcelError_FromComValue_PassesThroughOtherValues(object? value)
    {
        Assert.Equal(value, BatchExcel.Models.ExcelError.FromComValue(value));
    }
}
