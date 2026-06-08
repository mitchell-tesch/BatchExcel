using System.Runtime.InteropServices;

namespace BatchExcel.Services;

/// <summary>
/// COM IMessageFilter implementation that automatically retries calls rejected by Excel.
/// When Excel is busy (showing a dialog, recalculating, etc.) it rejects COM calls with
/// RPC_E_CALL_REJECTED. This filter catches those rejections and retries after a delay,
/// preventing transient failures in parallel Excel automation scenarios.
/// 
/// Must be registered on each STA thread that makes COM calls to Excel.
/// </summary>
public class ComMessageFilter : IOleMessageFilter
{
    private const int Cancel = -1;

    /// <summary>
    /// Registers this message filter on the current STA thread.
    /// Returns an IDisposable that will unregister the filter when disposed.
    /// </summary>
    public static IDisposable Register()
    {
        var newFilter = new ComMessageFilter();
        CoRegisterMessageFilter(newFilter, out var oldFilter);
        return new FilterRegistration(oldFilter);
    }

    // IOleMessageFilter implementation

    /// <summary>
    /// Called when an incoming COM call arrives while we're waiting.
    /// We handle all calls immediately.
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
        // SERVERCALL_ISHANDLED = 0 - we handle all incoming calls
        return 0;
    }

    /// <summary>
    /// Called when our outgoing COM call is rejected by the server (Excel).
    /// We retry after a brief delay instead of failing.
    /// </summary>
    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == 2) // SERVERCALL_RETRYLATER
        {
            // Retry after 200ms
            return 200;
        }

        // For SERVERCALL_REJECTED, retry if we haven't been waiting too long (< 30 seconds)
        if (dwTickCount < 30000)
        {
            // Retry after 500ms
            return 500;
        }

        // Give up after 30 seconds of retrying
        return Cancel;
    }

    /// <summary>
    /// Called when a message arrives while we're waiting for a COM response.
    /// We let the system handle it with PENDINGMSG_WAITDEFPROCESS.
    /// </summary>
    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        // PENDINGMSG_WAITDEFPROCESS = 2
        return 2;
    }

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        [MarshalAs(UnmanagedType.Interface)] IOleMessageFilter? lpMessageFilter,
        [MarshalAs(UnmanagedType.Interface)] out IOleMessageFilter? lplpMessageFilter);

    private class FilterRegistration : IDisposable
    {
        private readonly IOleMessageFilter? _oldFilter;
        private bool _disposed;

        public FilterRegistration(IOleMessageFilter? oldFilter)
        {
            _oldFilter = oldFilter;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                // Restore the previous filter (may legitimately be null on first registration —
                // the interop signature is nullable so we don't need the bang operator).
                CoRegisterMessageFilter(_oldFilter, out _);
                _disposed = true;
            }
        }
    }
}

/// <summary>
/// COM IOleMessageFilter interface definition for handling rejected calls.
/// </summary>
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

