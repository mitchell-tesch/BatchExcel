using System.IO;

namespace BatchExcel.Services;

/// <summary>
/// Fast filename sanitization using a pre-computed HashSet for character lookups.
/// </summary>
public static class FileNameSanitizer
{
    private static readonly HashSet<char> InvalidChars = new(Path.GetInvalidFileNameChars());

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
}

