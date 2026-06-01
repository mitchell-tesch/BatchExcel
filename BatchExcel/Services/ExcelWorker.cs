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

        // Caches of COM references
        Dictionary<string, dynamic>? sheetCache = null;
        (dynamic sheet, dynamic range, int offset)[]? inputRangeCache = null;
        dynamic[]? outputRangeCache = null;
        dynamic? workbook = null;

        try
        {
            excelApp = CreateExcelInstance();
            pid = ExcelProcessTracker.GetExcelProcessId(excelApp);
            if (pid != 0) ExcelProcessTracker.Track(pid);

            ctx.Log($"\t[Worker {ctx.WorkerId}] Excel started (PID: {pid})");

            workbook = excelApp.Workbooks.Open(ctx.CalculationPath, UpdateLinks: false);
            string calculationName = workbook.Name;

            SetCalculationMode(excelApp, XlCalculationManual, ctx.WorkerId, ctx.Log);

            sheetCache = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            inputRangeCache = BuildInputRangeCache(workbook, sheetCache);
            outputRangeCache = BuildOutputRangeCache(workbook, sheetCache);

            try
            {
                ProcessRunQueue(excelApp, workbook, calculationName, inputRangeCache, outputRangeCache);
                workbook.Close(SaveChanges: false);
            }
            catch
            {
                try { workbook.Close(SaveChanges: false); }
                catch
                {
                    // ignored
                }

                throw;
            }
            finally
            {
                workbook = null;
            }
        }
        catch (Exception ex)
        {
            ctx.Log($"\t[Worker {ctx.WorkerId}] FAILED: {ex.Message}");
        }
        finally
        {
            ctx.Log($"\t[Worker {ctx.WorkerId}] Shutting down...");

            if (excelApp != null)
                ExcelProcessTracker.SafeQuitExcel(excelApp, pid);
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
        string calculationName,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        while (ctx.RunQueue.TryDequeue(out var run))
        {
            if (ctx.CancellationToken.IsCancellationRequested)
            {
                ctx.Log($"\t[Worker {ctx.WorkerId}] Cancelled.");
                break;
            }

            try
            {
                ProcessSingleRunWithRetry(excelApp, workbook, calculationName, run, inputRangeCache, outputRangeCache);
            }
            catch (Exception ex)
            {
                ctx.Log($"\t[Worker {ctx.WorkerId}] ERROR on run '{run.Title}': {ex.Message}");
                // Mark the run as failed (Results stays null → CsvResultWriter reports "Failed").
                run.Results = null;
            }

            ctx.ReportRunCompleted();
        }
    }

    private void ProcessSingleRunWithRetry(
        dynamic excelApp,
        dynamic workbook,
        string calculationName,
        BatchRun run,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        const int maxAttempts = 3;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                ProcessSingleRun(excelApp, workbook, calculationName, run, inputRangeCache, outputRangeCache);
                return;
            }
            catch (COMException ex) when (attempt < maxAttempts && IsTransientComException(ex))
            {
                ctx.Log($"\t[Worker {ctx.WorkerId}] COM busy on '{run.Title}', retry {attempt}/{maxAttempts - 1}...");
                Thread.Sleep(1000 * attempt);
            }
            // Non-transient or final-attempt failures propagate to ProcessRunQueue's catch.
        }
    }

    private static bool IsTransientComException(COMException ex) =>
        ex.HResult == unchecked((int)0x80010001) || // RPC_E_CALL_REJECTED
        ex.HResult == unchecked((int)0x80010005) || // RPC_E_SERVERCALL_RETRYLATER
        ex.HResult == unchecked((int)0x800AC472);   // VBA_E_IGNORE (Excel busy)

    private void ProcessSingleRun(
        dynamic excelApp,
        dynamic workbook,
        string calculationName,
        BatchRun run,
        (dynamic sheet, dynamic range, int offset)[] inputRangeCache,
        dynamic[] outputRangeCache)
    {
        var totalRuns = ctx.Config.Calculations.Count;

        // Populate input fields using cached range references
        for (var i = 0; i < inputRangeCache.Length; i++)
        {
            var inputValue = run.Data[inputRangeCache[i].offset];
            // "*" means keep the existing calculation value
            if (inputValue is string s && s == "*") continue;
            inputRangeCache[i].range.Value = inputValue;
        }

        // Recalculate dirty cells only
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
        run.Results = results;

        // Save workbook copy if requested
        if (ctx.SaveRuns)
        {
            var runFileName = FileNameSanitizer.Sanitize($"{run.Index + 1}_{run.Title}_{calculationName}");
            var runFilePath = Path.Combine(ctx.OutFolder, runFileName);

            // SaveCopyAs writes a copy without rebinding the open workbook to the new path
            workbook.SaveCopyAs(runFilePath);

            if (ctx.PdfSheets.Count > 0)
            {
                var pdfPath = Path.ChangeExtension(runFilePath, ".pdf");
                PdfExporter.Export(workbook, ctx.PdfSheets, pdfPath);
            }
        }

        ctx.Log($"\t> ({run.Index + 1}/{totalRuns}) {run.Title}...done. [Worker {ctx.WorkerId}]");
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

