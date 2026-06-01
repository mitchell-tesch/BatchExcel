using System.Diagnostics;
using BatchExcel.Services;
using Xunit;

namespace BatchExcel.Tests;

public class ExcelProcessTrackerTests
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

    [Fact]
    public void Track_AddsToSet_IgnoringZero()
    {
        ExcelProcessTracker.Track(0);
        int killed = ExcelProcessTracker.KillAllTracked();
        Assert.Equal(0, killed);
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
        
        // Simulate WaitForExit(3000) returning false (still running)
        // In our mock, WaitForExit returns !RunningPids.Contains(pid).
        // To simulate timeout, we can make it always return false for this PID 
        // until Kill is called.
        
        ExcelProcessTracker.InteropProvider = mock;
        ExcelProcessTracker.Track(pid);
        
        // Mock dynamic excelApp
        // We can't easily mock 'dynamic' without a real COM object or ExpandoObject,
        // but SafeQuitExcel handles null/exceptions gracefully.
        ExcelProcessTracker.SafeQuitExcel(null, pid);

        // Should have checked if running, tried to wait, then killed
        Assert.Contains((int)pid, mock.KilledPids);
        Assert.Equal(0, ExcelProcessTracker.KillAllTracked()); // Should be untracked
    }
}
