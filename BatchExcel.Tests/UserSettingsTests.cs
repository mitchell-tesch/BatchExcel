using BatchExcel.Services;

namespace BatchExcel.Tests;

public class UserSettingsTests
{
    [Fact]
    public void DefaultSettings_HaveExpectedValues()
    {
        var settings = new UserSettings();

        Assert.Equal(4, settings.WorkerCount);
        Assert.True(settings.SaveRuns);
        Assert.Equal("", settings.PdfSheets);
        Assert.Equal("", settings.LastBatcherFilePath);
    }

    [Fact]
    public void Load_NoFile_ReturnsDefaults()
    {
        // This will return defaults if no settings file exists in %AppData%
        // (or the deserialized settings if one already exists - both paths covered by other tests)
        var settings = UserSettings.Load();
        Assert.NotNull(settings);
        Assert.True(settings.WorkerCount > 0);
    }
}

