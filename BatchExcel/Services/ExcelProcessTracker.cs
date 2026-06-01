using System.Diagnostics;
using System.Runtime.InteropServices;

namespace BatchExcel.Services;

/// <summary>
/// Abstraction for OS process interactions to enable deterministic testing.
/// </summary>
internal interface IProcessInterop
{
    void Kill(int pid);
    bool WaitForExit(int pid, int timeoutMs);
    bool IsRunning(int pid, string processName);
}

internal class DefaultProcessInterop : IProcessInterop
{
    public void Kill(int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            process.Kill();
        }
        catch (ArgumentException) { /* Already exited */ }
    }

    public bool WaitForExit(int pid, int timeoutMs)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            return process.WaitForExit(timeoutMs);
        }
        catch (ArgumentException) { return true; }
    }

    public bool IsRunning(int pid, string processName)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            return !process.HasExited && process.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException) { return false; }
    }
}

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

    internal static IProcessInterop InteropProvider { get; set; } = new DefaultProcessInterop();

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
        catch (Exception ex)
        {
            Debug.WriteLine($"[ExcelProcessTracker] Failed to resolve PID: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Registers an Excel process ID for tracking.
    /// </summary>
    public static void Track(uint pid)
    {
        if (pid == 0) return;
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
        if (pid == 0) return;
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
                if (InteropProvider.IsRunning((int)pid, "EXCEL"))
                {
                    InteropProvider.Kill((int)pid);
                    InteropProvider.WaitForExit((int)pid, 5000);
                    killed++;
                }
            }
            catch
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
        try
        {
            if (excelApp != null)
            {
                excelApp.DisplayAlerts = false;
                excelApp.Quit();
            }
        }
        catch
        {
            // Ignore quit errors
        }
        finally
        {
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
                    if (InteropProvider.IsRunning((int)pid, "EXCEL"))
                    {
                        // Wait briefly for graceful shutdown
                        if (!InteropProvider.WaitForExit((int)pid, 3000))
                        {
                            InteropProvider.Kill((int)pid);
                            InteropProvider.WaitForExit((int)pid, 5000);
                        }
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }

            Untrack(pid);
        }
    }
}
