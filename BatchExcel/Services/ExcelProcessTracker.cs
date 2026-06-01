using System.Diagnostics;
using System.Runtime.InteropServices;

namespace BatchExcel.Services;

/// <summary>
/// Tracks Excel process IDs to enable reliable cleanup of zombie processes.
/// Uses the Excel Application HWND to resolve the actual OS process ID.
/// </summary>
public static class ExcelProcessTracker
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private static readonly object Lock = new();
    private static readonly HashSet<uint> TrackedPids = new();

    /// <summary>
    /// Gets the OS process ID for an Excel Application COM object using its window handle.
    /// Returns 0 if the process ID cannot be resolved.
    /// </summary>
    public static uint GetExcelProcessId(dynamic excelApp)
    {
        try
        {
            IntPtr hwnd = new IntPtr((int)excelApp.Hwnd);
            GetWindowThreadProcessId(hwnd, out uint pid);
            return pid;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// Registers an Excel process ID for tracking.
    /// </summary>
    public static void Track(uint pid)
    {
        lock (Lock)
        {
            TrackedPids.Add(pid);
        }
    }

    /// <summary>
    /// Unregisters an Excel process ID (called after clean shutdown).
    /// </summary>
    public static void Untrack(uint pid)
    {
        lock (Lock)
        {
            TrackedPids.Remove(pid);
        }
    }

    /// <summary>
    /// Kills all tracked Excel processes that are still running.
    /// Call this on application exit or after unhandled exceptions.
    /// </summary>
    public static int KillAllTracked()
    {
        uint[] pids;
        lock (Lock)
        {
            pids = TrackedPids.ToArray();
            TrackedPids.Clear();
        }

        int killed = 0;
        foreach (var pid in pids)
        {
            try
            {
                using var process = Process.GetProcessById((int)pid);
                if (!process.HasExited && process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
                {
                    process.Kill();
                    process.WaitForExit(5000);
                    killed++;
                }
            }
            catch (ArgumentException)
            {
                // Process already exited
            }
            catch (Exception)
            {
                // Best effort cleanup
            }
        }

        return killed;
    }

    /// <summary>
    /// Safely quits an Excel Application and releases COM references.
    /// Falls back to process kill if graceful quit fails.
    /// </summary>
    public static void SafeQuitExcel(dynamic? excelApp, uint pid)
    {
        if (excelApp == null) return;

        try
        {
            excelApp.DisplayAlerts = false;
            excelApp.Quit();
        }
        catch
        {
            // Ignore quit errors
        }
        finally
        {
            // We do NOT call Marshal.FinalReleaseComObject(excelApp) here.
            // Since excelApp is passed as a 'dynamic', explicit release can cause 
            // InvalidComObjectException if the DLR has cached the reference.
            // We rely on the GC + Process.Kill fallback below for reliable cleanup.

            // Because we are relying entirely on the Garbage Collector to release RCWs,
            // we MUST use the "Double Tap" GC pattern. The first pass queues the finalizers
            // for root objects. The second pass cleans up transitively held COM objects 
            // (e.g., a Range held by a Sheet). Without the second pass, Excel will hang 
            // and trigger the Process.Kill fallback every time.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Verify the process actually exited, kill if not
            if (pid != 0)
            {
                try
                {
                    using var process = Process.GetProcessById((int)pid);
                    if (!process.HasExited)
                    {
                        // Wait briefly for graceful shutdown
                        if (!process.WaitForExit(3000))
                        {
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                    }
                }
                catch (ArgumentException)
                {
                    // Already exited - good
                }
            }

            Untrack(pid);
        }
    }
}

