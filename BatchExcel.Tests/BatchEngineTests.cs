using BatchExcel.Services;

namespace BatchExcel.Tests;

/// <summary>
/// Tests for <see cref="BatchEngine"/> behaviour that can be verified without launching
/// real Excel COM workers — Cancel/WasCancelled semantics, Dispose safety, etc.
/// End-to-end batch execution requires Excel and is covered by manual smoke-testing.
/// </summary>
public class BatchEngineTests
{
    [Fact]
    public void NewEngine_WasCancelled_IsFalse()
    {
        using var engine = new BatchEngine();
        Assert.False(engine.WasCancelled);
        Assert.Equal(0, engine.CompletedRunCount);
        Assert.Equal(0, engine.TotalIncludedRunCount);
    }

    [Fact]
    public void Cancel_SetsWasCancelled()
    {
        using var engine = new BatchEngine();
        engine.Cancel();
        Assert.True(engine.WasCancelled);
    }

    [Fact]
    public void Cancel_AfterDispose_DoesNotThrow()
    {
        var engine = new BatchEngine();
        engine.Dispose();

        // No exception expected — the engine must tolerate a UI-thread Cancel click that
        // races against the batch's own teardown (which Disposes the engine).
        var ex = Record.Exception(() => engine.Cancel());
        Assert.Null(ex);
        Assert.True(engine.WasCancelled);
    }

    [Fact]
    public void Cancel_MultipleTimes_IsIdempotent()
    {
        using var engine = new BatchEngine();
        engine.Cancel();
        engine.Cancel();
        engine.Cancel();
        Assert.True(engine.WasCancelled);
    }

    [Fact]
    public async Task RunAsync_NonexistentBatcherFile_ThrowsWithoutCrashing()
    {
        using var engine = new BatchEngine();
        var bogus = Path.Combine(Path.GetTempPath(), $"does_not_exist_{Guid.NewGuid():N}.xlsx");

        // ReadConfig opens the file via ClosedXML → FileNotFound surfaces as something IO-ish.
        await Assert.ThrowsAnyAsync<Exception>(() =>
            engine.RunAsync(bogus, workerCount: 1, saveRuns: false, pdfSheetsRaw: ""));

        // Engine should still be Cancellable + Disposable cleanly after the failure.
        engine.Cancel();
        Assert.True(engine.WasCancelled);
    }

    [Fact]
    public void Dispose_IsIdempotent()
    {
        var engine = new BatchEngine();
        engine.Dispose();
        // Second dispose must not throw.
        var ex = Record.Exception(() => engine.Dispose());
        Assert.Null(ex);
    }
}

