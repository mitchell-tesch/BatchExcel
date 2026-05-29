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
    /// Persists current settings to disk.
    /// </summary>
    public void Save()
    {
        try
        {
            Directory.CreateDirectory(SettingsDir);
            string json = JsonSerializer.Serialize(this, JsonOptions);
            File.WriteAllText(SettingsFile, json);
        }
        catch
        {
            // Best effort - don't crash the app if settings can't be saved
        }
    }
}

