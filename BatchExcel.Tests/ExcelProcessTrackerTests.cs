using System.Diagnostics;
using BatchExcel.Services;
using Xunit;

namespace BatchExcel.Tests;

// Use a collection to ensure these tests don't run in parallel with any other
// tests that might use the static ExcelProcessTracker.
[Collection("ExcelProcessTracker")]
public class ExcelProcessTrackerTests : IDisposable
{
    private class MockProcessInterop : IProcessInterop
    {
        public HashSet<int> RunningPids { get; } = new();
        public List<int> KilledPids { get; } = new();
        public List<(int pid, int timeout)> WaitCalls { get; } = new();

        public void Kill(int pid)
        {
            KilledPids.Add(pid);
            RunningPids.Remove(pid);
        }

        public bool WaitForExit(int pid, int timeoutMs)
        {
            WaitCalls.Add((pid, timeoutMs));
            return !RunningPids.Contains(pid);
        }

        public bool IsRunning(int pid, string processName)
        {
            return RunningPids.Contains(pid) && processName == "EXCEL";
        }
    }

    public void Dispose()
    {
        // Reset the provider and clear tracking state after each test
        ExcelProcessTracker.InteropProvider = new DefaultProcessInterop();
        ExcelProcessTracker.KillAllTracked();
    }

    [Fact]
    public void Track_AddsToSet_IgnoringZero()
    {
        var mock = new MockProcessInterop();
        ExcelProcessTracker.InteropProvider = mock;

        // PID 0 is ignored entirely (never tracked, never killed).
        ExcelProcessTracker.Track(0);

        // A non-zero PID *is* tracked: marking it as running in the mock means
        // KillAllTracked should pick it up and route a Kill through the provider.
        mock.RunningPids.Add(42);
        ExcelProcessTracker.Track(42);

        int killed = ExcelProcessTracker.KillAllTracked();

        Assert.Equal(1, killed);
        Assert.Contains(42, mock.KilledPids);
        Assert.DoesNotContain(0, mock.KilledPids);
    }

    [Fact]
    public void KillAllTracked_KillsOnlyRunningProcesses()
    {
        var mock = new MockProcessInterop();
        mock.RunningPids.Add(123);
        mock.RunningPids.Add(456);
        
        ExcelProcessTracker.InteropProvider = mock;
        ExcelProcessTracker.Track(123);
        ExcelProcessTracker.Track(456);
        ExcelProcessTracker.Track(789); // Not running according to mock

        int killed = ExcelProcessTracker.KillAllTracked();
        
        Assert.Equal(2, killed);
        Assert.Contains(123, mock.KilledPids);
        Assert.Contains(456, mock.KilledPids);
        Assert.DoesNotContain(789, mock.KilledPids);
    }

    [Fact]
    public void SafeQuitExcel_FallsBackToKill_IfProcessStaysRunning()
    {
        var mock = new MockProcessInterop();
        uint pid = 999;
        mock.RunningPids.Add((int)pid);
        
        ExcelProcessTracker.InteropProvider = mock;
        ExcelProcessTracker.Track(pid);
        
        // Even with a null excelApp, the finally block should still kill the PID.
        ExcelProcessTracker.SafeQuitExcel(null, pid);

        // Should have checked if running, tried to wait, then killed
        Assert.Contains((int)pid, mock.KilledPids);
        Assert.Equal(0, ExcelProcessTracker.KillAllTracked()); // Should be untracked
    }
}
