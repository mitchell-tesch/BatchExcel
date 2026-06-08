using BatchExcel.Services;

namespace BatchExcel.Tests;

/// <summary>
/// Tests for <see cref="IoRetry"/>. Uses synthetic <see cref="IOException"/>s with
/// hand-crafted HResults to exercise transient vs non-transient classification without
/// needing a real SMB share, AV scanner, etc.
/// </summary>
public class IoRetryTests
{
    // IOException HResults: 0x8007XXXX where the low word is the Win32 error code.
    private const int HrSharingViolation = unchecked((int)0x80070020); // ERROR_SHARING_VIOLATION
    private const int HrLockViolation    = unchecked((int)0x80070021); // ERROR_LOCK_VIOLATION
    private const int HrAccessDenied     = unchecked((int)0x80070005); // ERROR_ACCESS_DENIED
    private const int HrDiskFull         = unchecked((int)0x80070070); // ERROR_DISK_FULL (non-transient)
    private const int HrFileNotFound     = unchecked((int)0x80070002); // ERROR_FILE_NOT_FOUND (non-transient)

    [Fact]
    public void Run_NoException_ReturnsImmediately()
    {
        var attempts = 0;
        var result = IoRetry.Run(() => { attempts++; return 42; });

        Assert.Equal(42, result);
        Assert.Equal(1, attempts);
    }

    [Fact]
    public void Run_TransientSharingViolation_RetriesAndSucceeds()
    {
        var attempts = 0;
        var result = IoRetry.Run(() =>
        {
            attempts++;
            if (attempts < 3) throw new IOException("locked", HrSharingViolation);
            return "ok";
        }, maxAttempts: 5, baseDelayMs: 1); // tiny delay so test runs in ~ms

        Assert.Equal("ok", result);
        Assert.Equal(3, attempts);
    }

    [Fact]
    public void Run_TransientLockViolation_IsRetried()
    {
        var attempts = 0;
        var result = IoRetry.Run(() =>
        {
            attempts++;
            if (attempts < 2) throw new IOException("locked", HrLockViolation);
            return 1;
        }, maxAttempts: 3, baseDelayMs: 1);

        Assert.Equal(2, attempts);
        Assert.Equal(1, result);
    }

    [Fact]
    public void Run_TransientAccessDenied_IsRetried()
    {
        var attempts = 0;
        IoRetry.Run(() =>
        {
            attempts++;
            if (attempts < 2) throw new IOException("av busy", HrAccessDenied);
        }, maxAttempts: 3, baseDelayMs: 1);

        Assert.Equal(2, attempts);
    }

    [Fact]
    public void Run_NonTransientIOException_DoesNotRetry()
    {
        var attempts = 0;
        var ex = Assert.Throws<IOException>(() =>
        {
            IoRetry.Run(() =>
            {
                attempts++;
                throw new IOException("disk full", HrDiskFull);
            }, maxAttempts: 5, baseDelayMs: 1);
        });

        Assert.Equal(1, attempts);
        Assert.Equal(HrDiskFull, ex.HResult);
    }

    [Fact]
    public void Run_FileNotFound_DoesNotRetry()
    {
        var attempts = 0;
        Assert.Throws<IOException>(() =>
        {
            IoRetry.Run(() =>
            {
                attempts++;
                throw new IOException("missing", HrFileNotFound);
            }, maxAttempts: 5, baseDelayMs: 1);
        });

        Assert.Equal(1, attempts);
    }

    [Fact]
    public void Run_TransientExceptionExhaustsAttempts_PropagatesLast()
    {
        var attempts = 0;
        var ex = Assert.Throws<IOException>(() =>
        {
            IoRetry.Run(() =>
            {
                attempts++;
                throw new IOException("still locked", HrSharingViolation);
            }, maxAttempts: 3, baseDelayMs: 1);
        });

        Assert.Equal(3, attempts);
        Assert.Equal(HrSharingViolation, ex.HResult);
    }

    [Fact]
    public void Run_NonIOException_DoesNotRetry()
    {
        var attempts = 0;
        Assert.Throws<InvalidOperationException>(() =>
        {
            IoRetry.Run(() =>
            {
                attempts++;
                throw new InvalidOperationException();
            }, maxAttempts: 5, baseDelayMs: 1);
        });

        Assert.Equal(1, attempts);
    }

    [Fact]
    public void Run_VoidOverload_BehavesLikeReturning()
    {
        var attempts = 0;
        IoRetry.Run((Action)(() =>
        {
            attempts++;
            if (attempts < 2) throw new IOException("locked", HrSharingViolation);
        }), maxAttempts: 3, baseDelayMs: 1);

        Assert.Equal(2, attempts);
    }
}

