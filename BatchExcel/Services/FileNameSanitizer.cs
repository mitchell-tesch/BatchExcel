using System.IO;

namespace BatchExcel.Services;

/// <summary>
/// Fast filename sanitization using a pre-computed HashSet for character lookups.
/// </summary>
public static class FileNameSanitizer
{
    private static readonly HashSet<char> InvalidChars = new(Path.GetInvalidFileNameChars());

    /// <summary>
    /// Excel's <c>Workbooks.Open</c> / <c>Workbook.SaveCopyAs</c> / <c>ExportAsFixedFormat</c>
    /// cap the full path at ~218 chars regardless of Windows <c>LongPathsEnabled</c>.
    /// </summary>
    public const int ExcelMaxPathLength = 218;

    /// <summary>
    /// Replaces invalid filename characters with underscores in a single O(n) pass.
    /// </summary>
    public static string Sanitize(string fileName)
    {
        if (string.IsNullOrEmpty(fileName)) return fileName;

        return string.Create(fileName.Length, fileName, (span, src) =>
        {
            for (int i = 0; i < src.Length; i++)
                span[i] = InvalidChars.Contains(src[i]) ? '_' : src[i];
        });
    }

    /// <summary>
    /// Sanitizes and clamps the file name to <paramref name="maxLength"/> chars, preserving the
    /// extension. The stem is truncated; if the extension alone exceeds the budget the whole name
    /// is truncated from the right. Returns an empty string when <paramref name="maxLength"/> ≤ 0.
    /// </summary>
    public static string Sanitize(string fileName, int maxLength)
    {
        if (maxLength <= 0) return string.Empty;
        var sanitized = Sanitize(fileName);
        if (sanitized.Length <= maxLength) return sanitized;

        var ext = Path.GetExtension(sanitized);
        if (ext.Length >= maxLength)
            return sanitized[..maxLength];

        var stemBudget = maxLength - ext.Length;
        return string.Concat(sanitized.AsSpan(0, stemBudget), ext);
    }
}

