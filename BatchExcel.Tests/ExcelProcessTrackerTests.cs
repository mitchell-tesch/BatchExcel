using System.Diagnostics;
using BatchExcel.Services;
using Xunit;

namespace BatchExcel.Tests;

public class ExcelProcessTrackerTests
{
    [Fact]
    public void Track_AddsToSet()
    {
        uint pid = 999999; // Dummy PID
        ExcelProcessTracker.Track(pid);
        
        // No public way to check the set, but we can verify Untrack doesn't crash
        ExcelProcessTracker.Untrack(pid);
    }

    [Fact]
    public void Untrack_HandlesMissingPid()
    {
        uint pid = 888888;
        ExcelProcessTracker.Untrack(pid); // Should not throw
    }

    [Fact]
    public void KillAllTracked_HandlesDeadProcesses()
    {
        uint pid = 777777;
        ExcelProcessTracker.Track(pid);
        
        // This will try to get process by ID 777777, which likely doesn't exist.
        // It should catch the ArgumentException and proceed without throwing.
        int killed = ExcelProcessTracker.KillAllTracked();
        
        Assert.Equal(0, killed);
    }

    [Fact]
    public void KillAllTracked_ClearsSet()
    {
        uint pid = 666666;
        ExcelProcessTracker.Track(pid);
        ExcelProcessTracker.KillAllTracked();
        
        // Second call should definitely be 0
        int killed = ExcelProcessTracker.KillAllTracked();
        Assert.Equal(0, killed);
    }
}
