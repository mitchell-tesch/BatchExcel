using System.IO;

namespace BatchExcel.Services;

/// <summary>
/// Retries file IO that fails with transient sharing/lock violations — typical on network
/// shares where antivirus, the Windows Search Indexer, OneDrive, or SMB oplock breaks
/// briefly hold a file we just touched. Non-transient IOExceptions (and all other
/// exceptions) propagate immediately; the final-attempt failure also propagates so callers
/// can decide how to react.
/// </summary>
internal static class IoRetry
{
    private const int DefaultMaxAttempts = 5;
    private const int DefaultBaseDelayMs = 250;

    public static void Run(Action action, int maxAttempts = DefaultMaxAttempts, int baseDelayMs = DefaultBaseDelayMs)
    {
        Run(() => { action(); return 0; }, maxAttempts, baseDelayMs);
    }

    public static T Run<T>(Func<T> func, int maxAttempts = DefaultMaxAttempts, int baseDelayMs = DefaultBaseDelayMs)
    {
        for (var attempt = 1; ; attempt++)
        {
            try
            {
                return func();
            }
            catch (IOException ex) when (attempt < maxAttempts && IsTransient(ex))
            {
                Thread.Sleep(baseDelayMs * attempt); // linear backoff: 250, 500, 750, 1000 ms
            }
        }
    }

    private static bool IsTransient(IOException ex)
    {
        // IOException HResults are 0x8007XXXX where the low word is the Win32 error code.
        var win32 = ex.HResult & 0xFFFF;
        return win32 == 0x20    // ERROR_SHARING_VIOLATION
            || win32 == 0x21    // ERROR_LOCK_VIOLATION
            || win32 == 0x05;   // ERROR_ACCESS_DENIED (AV scanners sometimes surface this)
    }
}

