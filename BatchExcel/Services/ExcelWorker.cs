using System.Collections.Concurrent;
using System.IO;
using System.Runtime.InteropServices;
using BatchExcel.Models;

namespace BatchExcel.Services;

/// <summary>
/// Per-worker context bundling all state needed for a worker's lifetime.
/// Reduces parameter explosion when threading through the worker loop.
/// </summary>
internal sealed record WorkerContext(
    int WorkerId,
    string CalculationPath,
    string CalculationSourceName,
    BatchConfig Config,
    List<string> Macros,
    ConcurrentQueue<BatchRun> RunQueue,
    string OutFolder,
    bool SaveRuns,
    List<string> PdfSheets,
    CancellationToken CancellationToken,
    Action<string> Log,
    Action ReportRunCompleted);

/// <summary>
/// Single Excel worker that processes runs from a shared queue using one dedicated Excel instance.
/// Caches sheet and range COM references to minimize roundtrips.
/// </summary>
internal sealed class ExcelWorker(WorkerContext ctx)
{
    // Excel XlCalculation enum values
    private const int XlCalculationManual = -4135;

    /// <summary>
    /// Runs the worker loop. Must be invoked on an STA thread with a registered COM message filter.
    /// </summary>
    public void Run()
    {
        dynamic? excelApp = null;
        uint pid = 0;

        try
        {
            excelApp = CreateExcelInstance();
            pid = ExcelProcessTracker.GetExcelProcessId(excelApp);
            if (pid != 0) ExcelProcessTracker.Track(pid);

            ctx.Log($"\t[Worker {ctx.WorkerId}] Excel started (PID: {pid})");

            // Execute processing in a nested scope so local COM references 
            // naturally go out of scope before we trigger the GC collection.
            ExecuteRunLoop(excelApp);
        }
        catch (Exception ex)
        {
            ctx.Log($"\t[Worker {ctx.WorkerId}] FAILED: {ex.Message}");
        }
        finally
        {
            ctx.Log($"\t[Worker {ctx.WorkerId}] Shutting down...");

            if (excelApp != null)
            {
                var app = excelApp;
                excelApp = null; // Clear local reference before entering SafeQuitExcel
                ExcelProcessTracker.SafeQuitExcel(app, pid);
            }
        }
    }

    private void ExecuteRunLoop(dynamic excelApp)
    {
        dynamic? workbook = null;
        try
        {
            workbook = excelApp.Workbooks.Open(ctx.CalculationPath, UpdateLinks: false);

            SetCalculationMode(excelApp, XlCalculationManual, ctx.WorkerId, ctx.Log);

            var sheetCache = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            var inputRangeCache = BuildInputRangeCache(workbook, sheetCache);
            var outputRangeCache = BuildOutputRangeCache(workbook, sheetCache);

            try
            {
                ProcessRunQueue(excelApp, workbook, inputRangeCache, outputRangeCache);
                workbook.Close(SaveChanges: false);
            }
            catch
            {
                try { workbook.Close(SaveChanges: false); }
                catch { /* ignored */ }
                throw;
            }
        }
        finally
        {
            // The local variables 'workbook', 'sheetCache', etc. will go out of scope 
            // when this method returns, making them eligible for collection.
            workbook = null;
        }
    }

    private static dynamic CreateExcelInstance()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)
                      ?? throw new InvalidOperationException("Failed to create Excel instance.");

        app.Visible = false;
        app.DisplayAlerts = false;
        app.ScreenUpdating = false;
        app.EnableEvents = false;
        app.AskToUpdateLinks = false;
        app.Interactive = false;
        try { app.AutoRecover.Enabled = false; } catch { /* not always available */ }

        return app;
    }

    private (dynamic sheet, dynamic range, int offset)[] BuildInputRangeCache(
        dynamic workbook, Dictionary<string, dynamic> sheetCache)
    {
        var cache = new (dynamic, dynamic, int)[ctx.Config.InputFields.Count];
        for (var i = 0; i < ctx.Config.InputFields.Count; i++)
        {
            var field = ctx.Config.InputFields[i];
            var sheet = GetCachedSheet(workbook, field.Sheet, sheetCache);
            cache[i] = (sheet, sheet.Range[field.Range], field.ColumnOffset);
        }
        return cache;
    }

    private dynamic[] BuildOutputRangeCache(dynamic workbook, Dictionary<string, dynamic> sheetCache)
    {
        var cache = new dynamic[ctx.Config.OutputFields.Count];
        for (int i = 0; i < ctx.Config.OutputFields.Count; i++)
        {
            var field = ctx.Config.OutputFields[i];
            var sheet = GetCachedSheet(workbook, field.Sheet, sheetCache);
            cache[i] = sheet.Range[field.Range];
        }
        return cache;
    }

    private static dynamic GetCachedSheet(dynamic workbook, string sheetName, Dictionary<string, dynamic> cache)
    {
        if (!cache.TryGetValue(sheetName, out var sheet))
        {
            sheet = workbook.Sheets[sheetName];
            cache[sheetName] = sheet;
        }
        return sheet;
    }

    private void ProcessRunQueue(
        dynamic excelApp,
        dynamic workbook,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        // Check cancellation BEFORE dequeuing so a cancelled run isn't half-claimed:
        // the previous "dequeue then check" pattern would leave a dequeued run with
        // Results = null, which CsvResultWriter renders as "Failed" — misleading because
        // the run was never actually attempted. Leaving it in the queue keeps the run
        // available for another worker (if any) or, post-batch, it surfaces only via the
        // cancellation log line rather than as a spurious failure.
        while (!ctx.CancellationToken.IsCancellationRequested && ctx.RunQueue.TryDequeue(out var run))
        {
            try
            {
                ProcessSingleRunWithRetry(excelApp, workbook, run, inputRangeCache, outputRangeCache);
            }
            catch (Exception ex)
            {
                ctx.Log($"\t[Worker {ctx.WorkerId}] ERROR on run '{run.Title}': {ex.Message}");
                // Mark the run as failed (Results stays null → CsvResultWriter reports "Failed").
                run.Results = null;
            }

            ctx.ReportRunCompleted();
        }

        if (ctx.CancellationToken.IsCancellationRequested)
            ctx.Log($"\t[Worker {ctx.WorkerId}] Cancelled.");
    }

    private void ProcessSingleRunWithRetry(
        dynamic excelApp,
        dynamic workbook,
        BatchRun run,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        const int maxAttempts = 3;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                ProcessSingleRun(excelApp, workbook, run, inputRangeCache, outputRangeCache);
                return;
            }
            catch (COMException ex) when (IsTransientComException(ex))
            {
                // On the final attempt, propagate so ProcessRunQueue marks the run as Failed.
                // Without this explicit throw, exiting the loop normally would let the caller
                // believe the run succeeded (with stale/empty Results) — a silent-failure footgun.
                if (attempt == maxAttempts) throw;

                ctx.Log($"\t[Worker {ctx.WorkerId}] COM busy on '{run.Title}', retry {attempt}/{maxAttempts - 1}...");
                Thread.Sleep(1000 * attempt);
            }
            // Non-transient COMExceptions and any other exception propagate immediately.
        }
    }

    private static bool IsTransientComException(COMException ex) =>
        ex.HResult == unchecked((int)0x80010001) || // RPC_E_CALL_REJECTED
        ex.HResult == unchecked((int)0x80010005) || // RPC_E_SERVERCALL_RETRYLATER
        ex.HResult == unchecked((int)0x800AC472);   // VBA_E_IGNORE (Excel busy)

    private void ProcessSingleRun(
        dynamic excelApp,
        dynamic workbook,
        BatchRun run,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        // Populate input fields using cached range references
        for (var i = 0; i < inputRangeCache.Length; i++)
        {
            var inputValue = run.Data[inputRangeCache[i].offset];
            // "*" means keep the existing calculation value
            if (inputValue is string s && s == "*") continue;
            inputRangeCache[i].range.Value = inputValue;
        }

        // Recalculate dirty cells only.
        // NOTE: Workbook.Calculate doesn't exist in the Excel COM object model — only
        // Application.Calculate (all open workbooks), Worksheet.Calculate (one sheet) and
        // Range.Calculate (one range). Application.Calculate is safe here because each worker
        // owns its own Excel process and only ever has its single workbook open. If that ever
        // changes (e.g. a shared lookup workbook), switch to iterating workbook.Worksheets
        // and calling .Calculate on each sheet instead.
        excelApp.Calculate();

        // Run any configured macros
        foreach (var macroName in ctx.Macros)
        {
            excelApp.Run(macroName);
        }

        // Read output fields using cached range references
        var results = new object?[outputRangeCache.Length];
        for (var f = 0; f < outputRangeCache.Length; f++)
        {
            results[f] = outputRangeCache[f].Value;
        }
        // Assign Results BEFORE the optional save+PDF block. If SaveCopyAs / ExportAsFixedFormat
        // throws (disk full, locked, malformed PDF sheet name) we still want the calculated
        // values written to CSV — the save artifact is a convenience, not the canonical result.
        run.Results = results;

        // Save workbook copy if requested
        if (ctx.SaveRuns)
        {
            try
            {
                SaveRunArtifacts(workbook, run);
            }
            catch (Exception ex)
            {
                // Log but don't fail the run — Results is already populated, so CSV will show
                // "Completed" with valid values. Missing-PDF / missing-xlsx is a recoverable
                // problem that the user can re-export later from raw_output_fields.csv.
                ctx.Log($"\t[Worker {ctx.WorkerId}] WARNING: save artifacts for '{run.Title}' failed: {ex.Message}. " +
                        "Calculation results were retained.");
            }
        }

        ctx.Log($"\t> ({run.Index + 1}/{ctx.Config.Calculations.Count}) {run.Title}...done. [Worker {ctx.WorkerId}]");
    }

    private void SaveRunArtifacts(dynamic workbook, BatchRun run)
    {
        // Use the ORIGINAL calculation source filename (e.g. "calculation.xlsx") rather than
        // workbook.Name — which would be the per-worker copy's filename ("_worker_1_calculation.xlsx")
        // and would bleed worker-internal naming into the user-facing saved artifacts.
        //
        // Bonus: skipping workbook.Name saves one COM round-trip per run.
        //
        // Excel caps the full path at ~218 chars. outFolder is already preflighted by
        // BatchEngine, but a long run title can still push past the limit — clamp the file
        // name (preserving extension) so SaveCopyAs / ExportAsFixedFormat never throw on length.
        // ".pdf" and ".xlsx" are <= 5 chars, so reserving 5 for the extension swap is enough.
        var pathBudget = FileNameSanitizer.ExcelMaxPathLength - ctx.OutFolder.Length - 1;
        var runFileName = FileNameSanitizer.Sanitize(
            $"{run.Index + 1}_{run.Title}_{ctx.CalculationSourceName}",
            Math.Max(16, pathBudget));
        var runFilePath = Path.Combine(ctx.OutFolder, runFileName);

        // SaveCopyAs writes a copy without rebinding the open workbook to the new path
        workbook.SaveCopyAs(runFilePath);

        if (ctx.PdfSheets.Count > 0)
        {
            var pdfPath = Path.ChangeExtension(runFilePath, ".pdf");
            PdfExporter.Export(workbook, ctx.PdfSheets, pdfPath, ctx.Log);
        }
    }

    /// <summary>
    /// Safely sets the Excel Application calculation mode with retry.
    /// Excel can reject this call during certain states. Logs a warning if it never succeeds —
    /// running with the wrong calc mode silently can make a batch 10–100× slower.
    /// </summary>
    private static void SetCalculationMode(dynamic excelApp, int mode, int workerId, Action<string> log)
    {
        const int maxRetries = 3;
        for (var attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                excelApp.Calculation = mode;
                return;
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                log($"\t[Worker {workerId}] SetCalculationMode failed (attempt {attempt}/{maxRetries}): {ex.Message}");
                Thread.Sleep(250 * attempt);
            }
            catch (Exception ex)
            {
                log($"\t[Worker {workerId}] WARNING: SetCalculationMode failed after {maxRetries} attempts: {ex.Message}. " +
                    "Batch may run significantly slower if Excel stays in automatic calc mode.");
            }
        }
    }
}

