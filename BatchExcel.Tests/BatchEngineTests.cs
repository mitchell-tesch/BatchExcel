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

    // ---------------------------------------------------------------------
    // Disk-space preflight (pure helpers, no engine instance required)
    // ---------------------------------------------------------------------

    [Fact]
    public void EstimateRequiredDiskBytes_WorkerCopiesPlusOneStaged_PlusFixedOverhead()
    {
        // 4 workers + 1 transient staged file = 5 × source. SaveRuns/PDF disabled.
        const long source = 50L * 1024 * 1024; // 50 MB
        const long fixedOverhead = 10L * 1024 * 1024;

        var got = BatchEngine.EstimateRequiredDiskBytes(
            sourceSize: source,
            effectiveWorkers: 4,
            includedRunCount: 0,
            saveRuns: false,
            pdfSheetCount: 0);

        Assert.Equal(5 * source + fixedOverhead, got);
    }

    [Fact]
    public void EstimateRequiredDiskBytes_SaveRuns_AddsOneSourceSizePerIncludedRun()
    {
        const long source = 10L * 1024 * 1024;
        const int runs = 100;

        var withSave = BatchEngine.EstimateRequiredDiskBytes(source, 2, runs, saveRuns: true, pdfSheetCount: 0);
        var withoutSave = BatchEngine.EstimateRequiredDiskBytes(source, 2, runs, saveRuns: false, pdfSheetCount: 0);

        Assert.Equal(runs * source, withSave - withoutSave);
    }

    [Fact]
    public void EstimateRequiredDiskBytes_PdfSheets_AddsBudgetPerIncludedRun()
    {
        // PDF budget is the only difference between the two calls; verify it's non-zero
        // and scales linearly with run count.
        const long source = 4L * 1024 * 1024;
        var fewRuns = BatchEngine.EstimateRequiredDiskBytes(source, 2, 10, saveRuns: false, pdfSheetCount: 3);
        var manyRuns = BatchEngine.EstimateRequiredDiskBytes(source, 2, 100, saveRuns: false, pdfSheetCount: 3);

        Assert.True(manyRuns > fewRuns);
        // 10x runs → 10x PDF budget difference (worker copies are identical between calls).
        var perRunPdf = (manyRuns - fewRuns) / 90;
        Assert.True(perRunPdf > 0);
    }

    [Theory]
    [InlineData(-100L, 4, 0, false, 0)]   // negative source size clamped to 0
    [InlineData(1000L, 0, 0, false, 0)]   // zero workers clamped to 1
    [InlineData(1000L, 1, -5, true, 0)]   // negative run count clamped to 0
    public void EstimateRequiredDiskBytes_ClampsNegativeAndZeroInputs(
        long source, int workers, int runs, bool saveRuns, int pdfSheets)
    {
        var got = BatchEngine.EstimateRequiredDiskBytes(source, workers, runs, saveRuns, pdfSheets);
        // Result must always be at least the fixed overhead — never negative, never overflowing.
        Assert.True(got >= 10L * 1024 * 1024);
    }

    [Fact]
    public void EnsureSufficientDiskSpace_AvailableExceedsRequired_DoesNotThrow()
    {
        // 1 GB required, 10 GB available — comfortably above the 100 MB safety margin.
        var ex = Record.Exception(() => BatchEngine.EnsureSufficientDiskSpace(
            requiredBytes: 1_000_000_000L,
            availableBytes: 10_000_000_000L,
            outFolder: @"C:\out"));
        Assert.Null(ex);
    }

    [Fact]
    public void EnsureSufficientDiskSpace_AvailableBelowRequired_Throws()
    {
        var ex = Assert.Throws<InsufficientDiskSpaceException>(() => BatchEngine.EnsureSufficientDiskSpace(
            requiredBytes: 10_000_000_000L,
            availableBytes: 1_000_000_000L,
            outFolder: @"C:\out"));

        // Message must name the folder and quote both numbers in MB so the user can act on it.
        Assert.Contains("out", ex.Message);
        Assert.Contains("MB", ex.Message);
    }

    [Fact]
    public void EnsureSufficientDiskSpace_TightSqueeze_StillThrowsBecauseOfSafetyMargin()
    {
        // Available is exactly the required amount — falls inside the 100 MB safety margin
        // and should throw rather than skating by on a no-headroom pass.
        const long need = 5L * 1024 * 1024 * 1024; // 5 GB
        Assert.Throws<InsufficientDiskSpaceException>(() => BatchEngine.EnsureSufficientDiskSpace(
            requiredBytes: need,
            availableBytes: need,
            outFolder: @"C:\out"));
    }
}

