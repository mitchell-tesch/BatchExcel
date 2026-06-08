using System.Collections.Concurrent;
using System.IO;
using System.Text;
using BatchExcel.Models;

namespace BatchExcel.Services;

/// <summary>
/// Orchestrates parallel batch processing of Excel calculations.
/// Delegates per-worker COM logic to <see cref="ExcelWorker"/> and file I/O to
/// <see cref="BatcherReader"/>, <see cref="CalculationHeaderWriter"/>, and <see cref="CsvResultWriter"/>.
/// </summary>
public class BatchEngine : IDisposable
{
    private const string LogFileName = "batch_log.log";

    public event Action<string>? LogMessage;
    public event Action<int, int>? ProgressChanged; // (completed, total)

    private int _completedRuns;
    private int _totalIncludedRuns;
    private long _lastProgressUpdateTicks;
    private readonly CancellationTokenSource _cts = new();

    /// <summary>
    /// True if the batch was cancelled mid-run. Set when <see cref="Cancel"/> is invoked.
    /// The VM uses this to distinguish a clean completion from a user-cancelled partial run.
    /// </summary>
    public bool WasCancelled => _wasCancelled;
    private volatile bool _wasCancelled;

    /// <summary>Number of runs that actually completed (success or failure) before the engine returned.</summary>
    public int CompletedRunCount => _completedRuns;

    /// <summary>Total number of runs originally scheduled (included in the batch).</summary>
    public int TotalIncludedRunCount => _totalIncludedRuns;

    // Log-to-file state. Messages emitted before the output folder exists are buffered;
    // once OpenLogFile() is called, the buffer is flushed and subsequent messages stream directly.
    private readonly StringBuilder _earlyLogBuffer = new();
    private readonly Lock _logFileLock = new();
    private StreamWriter? _logFileWriter;

    /// <summary>
    /// Executes the full batch workflow: read config, write headers, run batch, write results.
    /// </summary>
    public async Task RunAsync(string batcherFilePath, int workerCount, bool saveRuns, string pdfSheetsRaw)
    {
        var batchStart = DateTime.Now;
        var timestamp = batchStart.ToString("yyMMdd-HHmmss");

        var pdfSheets = ParseCsvList(pdfSheetsRaw);

        Log("BatchExcel");
        Log("Parallel Processor Edition");
        Log("--------------------------");

        // Step 1: Read configuration (direct XML — no interop)
        Log("\nReading batching input...");
        var config = BatcherReader.ReadConfig(batcherFilePath);

        // IncludedCalcCount enumerates Calculations on each access — cache once for the
        // multiple comparisons / log lines below to avoid O(N) scans on large run counts.
        var includedCount = config.IncludedCalcCount;
        var totalCount = config.Calculations.Count;

        Log($"\t> Calculation file: {config.CalculationFile}");
        Log($"\t> No. of input fields: {config.InputFields.Count}");
        Log($"\t> No. of output fields: {config.OutputFields.Count}");
        Log($"\t> No. of skipped fields: {config.SkipFields.Count}");
        Log($"\t> No. of calculations: {totalCount} ({includedCount} included and {totalCount - includedCount} skipped)");

        if (includedCount == 0)
        {
            Log("\nNo calculations to process.");
            return;
        }

        // Resolve calculation path relative to batcher file
        var batcherDir = Path.GetDirectoryName(Path.GetFullPath(batcherFilePath))!;
        var calculationFullPath = Path.GetFullPath(Path.Combine(batcherDir, config.CalculationFile));

        if (!File.Exists(calculationFullPath))
            throw new FileNotFoundException($"Calculation spreadsheet not found: {calculationFullPath}");

        // Step 2: Dry-run validation — fail fast before spinning up Excel workers if any
        // input/output field references a missing sheet or unresolvable range.
        Log("\nValidating calculation spreadsheet... ");
        CalculationValidator.Validate(calculationFullPath, config);
        Log("done.");

        // Step 3: Create output folder
        var outFolder = Path.Combine(batcherDir, $"batch_run_{timestamp}");
        Directory.CreateDirectory(outFolder);
        Log($"Output folder: {outFolder}");

        // Open log file in output folder. Buffered early log messages will be flushed automatically.
        OpenLogFile(outFolder);

        // Step 4: Clamp worker count.
        var requestedWorkers = workerCount;
        var effectiveWorkers = Math.Clamp(requestedWorkers, 1, Environment.ProcessorCount * 2);
        effectiveWorkers = Math.Min(effectiveWorkers, includedCount);
        if (effectiveWorkers != requestedWorkers)
            Log($"Worker count adjusted from {requestedWorkers} to {effectiveWorkers} (bounded by CPU and run count).");

        // Step 4b: Preflight path lengths. Excel's COM APIs cap full paths at ~218 chars regardless
        // of Windows LongPathsEnabled, and PathTooLongException is not transient so it would abort
        // the batch mid-flight with a raw stack trace. Fail fast here with an actionable message.
        var sourceName = Path.GetFileName(calculationFullPath);
        // outFolder + "\" + "_worker_" + <digits> + "_" + sourceName
        var workerCopyLen = outFolder.Length + 1 + "_worker_".Length
                            + effectiveWorkers.ToString().Length + 1 + sourceName.Length;
        if (workerCopyLen > FileNameSanitizer.ExcelMaxPathLength)
        {
            var msg = $"Output paths would exceed Excel's {FileNameSanitizer.ExcelMaxPathLength}-char limit "
                      + $"(worker copy path = {workerCopyLen}). "
                      + $"Move the batcher closer to the drive root or shorten the calculation filename "
                      + $"('{sourceName}', {sourceName.Length} chars).";
            Log("ERROR: " + msg);
            throw new PathTooLongException(msg);
        }

        // Step 5: Create a single staging copy of the calculation in the output folder, write the
        // headers into it once via OpenXML, then duplicate it for each worker. Header inputs are
        // identical across workers, so writing once and copying is much cheaper than opening and
        // rewriting every worker copy. The user's original calculation file is never mutated.
        Log("\nConfiguring worker calculation copies... ");
        var workerCalcPaths = CreateWorkerCalcCopiesFromStaged(
            effectiveWorkers, calculationFullPath, outFolder, config.HeaderInputs, batchStart);
        Log("done.");

        // Step 6: Parallel batch processing
        Log($"\nStarting batching with {effectiveWorkers} parallel worker(s)...");

        _completedRuns = 0;
        _totalIncludedRuns = includedCount;
        _lastProgressUpdateTicks = 0;
        ProgressChanged?.Invoke(0, _totalIncludedRuns);

        var runQueue = new ConcurrentQueue<BatchRun>(config.Calculations.Where(r => r.Include));
        var macros = config.Macros; // cache once - avoids reparsing CSV per run

        // Log skipped runs
        foreach (var run in config.Calculations.Where(r => !r.Include))
        {
            Log($"\t> ({run.Index + 1}/{totalCount}) {run.Title} - skipped.");
        }

        try
        {
            await LaunchAndAwaitWorkers(effectiveWorkers, workerCalcPaths, config, macros, runQueue, outFolder, saveRuns, pdfSheets);
        }
        finally
        {
            // Ensure final progress is reported
            ProgressChanged?.Invoke(_completedRuns, _totalIncludedRuns);
            CleanupWorkerCalculationCopies(workerCalcPaths);
        }

        // If every worker failed to start (e.g. workbook.Open threw for all of them) the queue
        // will still contain undequeued runs. Without this check the engine would happily claim
        // "Batch completed successfully" while producing zero results.
        // Skip the check when the user cancelled — undequeued items in that case are expected,
        // not a failure mode, and shouldn't be reported as "every worker failed to start".
        if (!_wasCancelled && !runQueue.IsEmpty)
        {
            var undone = runQueue.Count;
            throw new InvalidOperationException(
                $"Batch aborted: {undone} run(s) were never processed because every worker failed to start. " +
                "See the log above for the per-worker failure reason.");
        }

        // Step 7: Write results back (direct file access — no Excel needed).
        // Always write whatever results we have, even on cancellation — partial results are
        // more useful to the user than no results.
        Log("\nWriting results... ");
        bool originalUpdated = BatcherReader.WriteResults(batcherFilePath, config, outFolder);
        if (!originalUpdated)
            Log($"WARNING: could not write results to '{batcherFilePath}' (file in use?). " +
                $"Results were saved to the copy in the output folder.");

        // Step 8: Write CSV
        CsvResultWriter.Write(outFolder, config);
        Log("done.");

        var elapsed = DateTime.Now - batchStart;
        if (_wasCancelled)
            Log($"\nBatch cancelled after {elapsed.TotalSeconds:F1} seconds ({_completedRuns}/{_totalIncludedRuns} runs completed before cancel).");
        else
            Log($"\nBatch completed successfully in {elapsed.TotalSeconds:F1} seconds.");
    }

    /// <summary>Cancels the running batch operation. Safe to call after Dispose.</summary>
    public void Cancel()
    {
        Log("\nCancellation requested...");
        _wasCancelled = true;
        try
        {
            // _cts.Cancel() throws ObjectDisposedException if Dispose() has already run —
            // possible if the UI thread races a Cancel click against the batch's finally block.
            if (!_cts.IsCancellationRequested)
                _cts.Cancel();
        }
        catch (ObjectDisposedException)
        {
            // Batch already finished or torn down — nothing to cancel.
        }
    }

    public void Dispose()
    {
        CloseLogFile();
        _cts.Dispose();
        GC.SuppressFinalize(this);
    }

    private async Task LaunchAndAwaitWorkers(
        int workerCount,
        List<string> workerCalculationPaths,
        BatchConfig config,
        List<string> macros,
        ConcurrentQueue<BatchRun> runQueue,
        string outFolder,
        bool saveRuns,
        List<string> pdfSheets)
    {
        var workerTasks = new List<Task>();

        for (var i = 0; i < workerCount; i++)
        {
            var workerId = i + 1;
            var workerCalcPath = workerCalculationPaths[i];
            var staggerMs = i * 200; // each worker waits this long before its STA thread starts

            var ctx = new WorkerContext(
                WorkerId: workerId,
                CalculationPath: workerCalcPath,
                Config: config,
                Macros: macros,
                RunQueue: runQueue,
                OutFolder: outFolder,
                SaveRuns: saveRuns,
                PdfSheets: pdfSheets,
                CancellationToken: _cts.Token,
                Log: Log,
                ReportRunCompleted: OnRunCompleted);

            // Stagger asynchronously so we don't burn a dedicated STA thread just to Sleep on it.
            workerTasks.Add(StaggeredWorker(staggerMs, ctx));
        }

        await Task.WhenAll(workerTasks);
        return;

        static async Task StaggeredWorker(int delayMs, WorkerContext ctx)
        {
            if (delayMs > 0)
                await Task.Delay(delayMs).ConfigureAwait(false);
            await RunOnStaThread(() => new ExcelWorker(ctx).Run()).ConfigureAwait(false);
        }
    }

    private void OnRunCompleted()
    {
        var completed = Interlocked.Increment(ref _completedRuns);
        ReportProgressThrottled(completed, _totalIncludedRuns);
    }

    private static List<string> CreateWorkerCalcCopiesFromStaged(
        int workerCount,
        string calculationFullPath,
        string outFolder,
        Dictionary<string, object?> headerInputs,
        DateTime batchStart)
    {
        // Stage: copy source → staging file, write headers once.
        var fileName = Path.GetFileName(calculationFullPath);
        var stagedPath = Path.Combine(outFolder, $"_staged_{fileName}");
        // Wrap copies in IoRetry — source and destination may both be on an SMB share where
        // AV / Search Indexer / OneDrive briefly hold freshly-written files.
        IoRetry.Run(() => File.Copy(calculationFullPath, stagedPath, overwrite: true));

        try
        {
            CalculationHeaderWriter.Write(stagedPath, headerInputs, batchStart);

            // Fan out: duplicate the staged (header-stamped) file to each worker copy in parallel.
            // File.Copy from one source to many destinations parallelises well on SSDs; the
            // small DOP cap keeps HDD-bound systems from thrashing.
            var paths = new string[workerCount];
            var copyDop = Math.Min(workerCount, Math.Max(2, Environment.ProcessorCount / 2));
            Parallel.For(0, workerCount, new ParallelOptions { MaxDegreeOfParallelism = copyDop }, i =>
            {
                var copy = Path.Combine(outFolder, $"_worker_{i + 1}_{fileName}");
                IoRetry.Run(() => File.Copy(stagedPath, copy, overwrite: true));
                paths[i] = copy;
            });
            return paths.ToList();
        }
        finally
        {
            // Wrap delete in IoRetry — on SMB shares or AV-active workstations the staged
            // file's handle may still be held briefly after CalculationHeaderWriter.Write
            // returns. A failed delete here would leak a "_staged_*.xlsx" file into the
            // output folder, which is cosmetic but confusing.
            try { IoRetry.Run(() => File.Delete(stagedPath)); } catch { /* best effort */ }
        }
    }

    private static void CleanupWorkerCalculationCopies(List<string> paths)
    {
        foreach (var p in paths)
        {
            // Same SMB/AV rationale as the staged file delete above.
            try { IoRetry.Run(() => File.Delete(p)); } catch { /* best effort */ }
        }
    }

    private static List<string> ParseCsvList(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return [];
        return raw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
    }

    /// <summary>
    /// Reports progress at most every 100ms to avoid flooding the UI dispatcher.
    /// </summary>
    private void ReportProgressThrottled(int completed, int total)
    {
        var now = Environment.TickCount64;
        var last = Interlocked.Read(ref _lastProgressUpdateTicks);

        if (completed < total && now - last < 100) return;
        if (Interlocked.CompareExchange(ref _lastProgressUpdateTicks, now, last) == last)
        {
            ProgressChanged?.Invoke(completed, total);
        }
    }

    private void Log(string message)
    {
        LogMessage?.Invoke(message);
        WriteToLogFile(message);
    }

    /// <summary>
    /// Writes a log line to the file writer if open, otherwise buffers it for later flushing.
    /// </summary>
    private void WriteToLogFile(string message)
    {
        lock (_logFileLock)
        {
            if (_logFileWriter != null)
            {
                try { _logFileWriter.WriteLine(message); } catch { /* best effort */ }
            }
            else
            {
                _earlyLogBuffer.AppendLine(message);
            }
        }
    }

    /// <summary>
    /// Opens the batch log file in the output folder and flushes any buffered early messages.
    /// </summary>
    private void OpenLogFile(string outFolder)
    {
        lock (_logFileLock)
        {
            try
            {
                var logPath = Path.Combine(outFolder, LogFileName);
                _logFileWriter = new StreamWriter(logPath, append: false) { AutoFlush = true };

                if (_earlyLogBuffer.Length <= 0) return;
                _logFileWriter.Write(_earlyLogBuffer.ToString());
                _earlyLogBuffer.Clear();
            }
            catch
            {
                _logFileWriter = null; // logging to file is best-effort
            }
        }
    }

    private void CloseLogFile()
    {
        lock (_logFileLock)
        {
            try { _logFileWriter?.Dispose(); }
            catch
            {
                // ignored
            }

            _logFileWriter = null;
        }
    }

    /// <summary>
    /// Runs an action on a new STA thread (required for Excel COM) with a registered COM message filter.
    /// </summary>
    private static Task RunOnStaThread(Action action)
    {
        var tcs = new TaskCompletionSource();

        var thread = new Thread(() =>
        {
            using var filter = ComMessageFilter.Register();
            try
            {
                action();
                tcs.SetResult();
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        return tcs.Task;
    }
}
