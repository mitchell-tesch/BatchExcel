using System.Runtime.InteropServices;

namespace BatchExcel.Services;

internal static class CpuInfo
{
    private const int RelationProcessorCore = 0;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetLogicalProcessorInformationEx(int relationshipType, IntPtr buffer, ref uint returnedLength);

    /// <summary>Number of physical CPU cores, or the logical processor count if it can't be determined.</summary>
    public static int PhysicalCoreCount()
    {
        uint length = 0;
        GetLogicalProcessorInformationEx(RelationProcessorCore, IntPtr.Zero, ref length);
        if (length == 0)
            return Environment.ProcessorCount;

        var buffer = Marshal.AllocHGlobal((int)length);
        try
        {
            if (!GetLogicalProcessorInformationEx(RelationProcessorCore, buffer, ref length))
                return Environment.ProcessorCount;

            // One variable-length SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX record per core; Size is at offset 4.
            var cores = 0;
            for (var offset = 0; offset < length; cores++)
            {
                var size = Marshal.ReadInt32(buffer, offset + 4);
                if (size <= 0) break;
                offset += size;
            }
            return cores > 0 ? cores : Environment.ProcessorCount;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }
}
