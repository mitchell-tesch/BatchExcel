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
    /// </summary>
    public static uint GetExcelProcessId(dynamic excelApp)
    {
        IntPtr hwnd = new IntPtr((int)excelApp.Hwnd);
        GetWindowThreadProcessId(hwnd, out uint pid);
        return pid;
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
            try
            {
                Marshal.FinalReleaseComObject(excelApp);
            }
            catch
            {
                // Ignore
            }

            // FinalReleaseComObject above has already driven the RCW refcount to zero. A single
            // GC + finalizer wait is enough to flush any transitively-held RCWs; the previous
            // double-collect pattern was redundant once we stopped relying on the GC alone.
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

