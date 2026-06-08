using System.IO;
using System.Text.Json;

namespace BatchExcel.Services;

/// <summary>
/// Persists user preferences to %AppData%\BatchExcel\settings.json so they survive between sessions.
/// </summary>
public class UserSettings
{
    private static readonly string SettingsDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BatchExcel");

    private static readonly string SettingsFile = Path.Combine(SettingsDir, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public int WorkerCount { get; set; } = 4;
    public bool SaveRuns { get; set; } = true;
    public string PdfSheets { get; set; } = "";
    public string LastBatcherFilePath { get; set; } = "";

    /// <summary>
    /// Loads settings from disk, returning defaults if the file doesn't exist or is corrupted.
    /// </summary>
    public static UserSettings Load()
    {
        try
        {
            if (File.Exists(SettingsFile))
            {
                string json = File.ReadAllText(SettingsFile);
                var settings = JsonSerializer.Deserialize<UserSettings>(json);
                if (settings != null) return settings;
            }
        }
        catch
        {
            // Fall through to defaults on any error (corrupt file, IO error, etc.)
        }
        return new UserSettings();
    }

    /// <summary>
    /// Persists current settings to disk atomically: writes to a tempfile in the same
    /// directory, then moves over the target. Prevents a crash mid-write from leaving
    /// the settings file empty/corrupt (which would silently reset to defaults next launch).
    /// </summary>
    public void Save()
    {
        try
        {
            Directory.CreateDirectory(SettingsDir);
            string json = JsonSerializer.Serialize(this, JsonOptions);

            // Same-directory tempfile so File.Move is an atomic rename on NTFS rather than a copy.
            var tempPath = SettingsFile + ".tmp";
            File.WriteAllText(tempPath, json);
            File.Move(tempPath, SettingsFile, overwrite: true);
        }
        catch
        {
            // Best effort - don't crash the app if settings can't be saved
        }
    }
}

