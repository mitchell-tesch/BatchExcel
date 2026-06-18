using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using BatchExcel.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;

namespace BatchExcel.ViewModels;

public partial class MainViewModel : ObservableObject
{
    // Cap in-memory log to ~256KB to keep the TextBox responsive on long batches.
    // Full unbounded log is still written to batch_log.log on disk.
    private const int MaxLogChars = 256 * 1024;
    private const string LogTruncationNotice = "[... earlier log lines truncated; see batch_log.log for full output ...]\n";

    private BatchEngine? _engine;
    private readonly StringBuilder _logBuilder = new();
    private readonly object _logLock = new();
    private readonly UserSettings _settings;
    private DispatcherTimer? _logFlushTimer;
    private DispatcherTimer? _settingsSaveTimer;
    private bool _logDirty;

    public MainViewModel()
    {
        _settings = UserSettings.Load();
        _batcherFilePath = _settings.LastBatcherFilePath;
        _workerCount = _settings.WorkerCount;
        _saveRuns = _settings.SaveRuns;
        _pdfSheets = _settings.PdfSheets;
    }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunBatchCommand))]
    private string _batcherFilePath;

    [ObservableProperty]
    private int _workerCount;

    [ObservableProperty]
    private bool _saveRuns;

    [ObservableProperty]
    private string _pdfSheets;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunBatchCommand))]
    [NotifyCanExecuteChangedFor(nameof(CancelBatchCommand))]
    private bool _isRunning;

    [ObservableProperty]
    private double _progressPercent;

    [ObservableProperty]
    private string _progressText = "Ready";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveLogAsCommand))]
    private string _logOutput = "";

    /// <summary>True after a batch failure — drives the red status text + snackbar.</summary>
    [ObservableProperty]
    private bool _hasError;

    /// <summary>
    /// Severity flag for <see cref="NotificationRequested"/>. View maps these to
    /// WPF-UI <c>ControlAppearance</c> + icon without leaking Wpf.Ui types into the VM.
    /// </summary>
    public enum NotificationKind { Success, Warning, Error }

    /// <summary>Raised when the VM wants a transient Fluent snackbar shown.</summary>
    public event Action<string, string, NotificationKind>? NotificationRequested;

    // Persist settings whenever a user-editable property changes (debounced to avoid
    // a disk write on every keystroke when bindings use UpdateSourceTrigger=PropertyChanged).
    partial void OnBatcherFilePathChanged(string value) => SchedulePersistSettings();
    partial void OnWorkerCountChanged(int value) => SchedulePersistSettings();
    partial void OnSaveRunsChanged(bool value) => SchedulePersistSettings();
    partial void OnPdfSheetsChanged(string value) => SchedulePersistSettings();

    private void SchedulePersistSettings()
    {
        // If we're not on a UI thread (no Application), persist immediately.
        if (Application.Current?.Dispatcher is not { } dispatcher)
        {
            PersistSettings();
            return;
        }

        dispatcher.Invoke(() =>
        {
            _settingsSaveTimer ??= new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _settingsSaveTimer.Tick -= OnSettingsSaveTick;
            _settingsSaveTimer.Tick += OnSettingsSaveTick;
            _settingsSaveTimer.Stop();
            _settingsSaveTimer.Start();
        });
    }

    private void OnSettingsSaveTick(object? sender, EventArgs e)
    {
        _settingsSaveTimer?.Stop();
        PersistSettings();
    }

    private void PersistSettings()
    {
        _settings.LastBatcherFilePath = BatcherFilePath;
        _settings.WorkerCount = WorkerCount;
        _settings.SaveRuns = SaveRuns;
        _settings.PdfSheets = PdfSheets;
        _settings.Save();
    }

    [RelayCommand]
    private void BrowseFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Select BatchExcel Workbook",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog() == true)
        {
            BatcherFilePath = dialog.FileName;
        }
    }

    [RelayCommand]
    private void NewFromBatcherTemplate()
    {
        // Locate the bundled batcher template
        var batcherTemplateSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "BatchExcel_template.xlsx");
        if (!File.Exists(batcherTemplateSource))
        {
            MessageBox.Show("Batcher template file not found. Ensure BatchExcel_template.xlsx is in the Resources folder.",
                "Batcher Template Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Let user choose where to save the new batcher workbook
        var dialog = new SaveFileDialog
        {
            Title = "Create New Batcher Workbook",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "BatchExcel_batcher.xlsx",
        };

        if (dialog.ShowDialog() != true) return;
        try
        {
            File.Copy(batcherTemplateSource, dialog.FileName, overwrite: true);
            BatcherFilePath = dialog.FileName;

            // Open in default application (Excel)
            Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AppendLog($"Created new batcher from template: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to create batcher: {ex.Message}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void OpenBatcherInExcel()
    {
        if (string.IsNullOrWhiteSpace(BatcherFilePath) || !File.Exists(BatcherFilePath))
        {
            MessageBox.Show("No batcher workbook selected or file does not exist.",
                "Cannot Open", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(BatcherFilePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to open file: {ex.Message}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand(CanExecute = nameof(CanSaveLog))]
    private void SaveLogAs()
    {
        var dialog = new SaveFileDialog
        {
            Title = "Save Log As",
            Filter = "Log File (*.log)|*.log|Text File (*.txt)|*.txt|All files (*.*)|*.*",
            FileName = $"BatchExcel_log_{DateTime.Now:yyMMdd-HHmmss}.log",
        };

        if (dialog.ShowDialog() != true) return;
        try
        {
            File.WriteAllText(dialog.FileName, LogOutput);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save log: {ex.Message}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private bool CanSaveLog() => !string.IsNullOrEmpty(LogOutput);

    [RelayCommand(CanExecute = nameof(CanRunBatch))]
    private async Task RunBatch()
    {
        if (string.IsNullOrWhiteSpace(BatcherFilePath) || !File.Exists(BatcherFilePath))
        {
            AppendLog("ERROR: Please select a valid batcher workbook file.");
            FlushLog();
            return;
        }

        IsRunning = true;
        ProgressPercent = 0;
        ProgressText = "Starting...";
        HasError = false;
        lock (_logLock)
        {
            _logBuilder.Clear();
            _logDirty = true;
        }
        LogOutput = "";

        // Start log flush timer (UI thread) - batches log updates at 100ms intervals
        _logFlushTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(100)
        };
        _logFlushTimer.Tick += (_, _) => FlushLog();
        _logFlushTimer.Start();

        _engine = new BatchEngine();
        _engine.LogMessage += OnLogMessage;
        _engine.ProgressChanged += OnProgressChanged;

        try
        {
            // RunAsync's first await is deep inside the worker loop — the prep work (config
            // read, validation, file copies, header writes) runs synchronously up to that point.
            // Offload to a thread-pool thread so the UI stays responsive during prep on large
            // batchers / slow disks. Log + Progress callbacks already marshal back to the UI
            // thread (via thread-safe buffer + dispatcher BeginInvoke).
            await Task.Run(() => _engine.RunAsync(BatcherFilePath, WorkerCount, SaveRuns, PdfSheets));

            if (_engine.WasCancelled)
            {
                ProgressText = $"Cancelled ({_engine.CompletedRunCount}/{_engine.TotalIncludedRunCount} done)";
                NotificationRequested?.Invoke(
                    "Batch cancelled",
                    $"{_engine.CompletedRunCount} of {_engine.TotalIncludedRunCount} runs completed before cancel. " +
                    "Partial results were written to the output folder.",
                    NotificationKind.Warning);
            }
            else
            {
                ProgressText = "Complete";
                ProgressPercent = 100;
                NotificationRequested?.Invoke(
                    "Batch complete",
                    "All runs finished — see log for details.",
                    NotificationKind.Success);
            }
        }
        catch (PathTooLongException ex)
        {
            // Preflight in BatchEngine.RunAsync throws this with an actionable message.
            // Surface it as a Fluent snackbar (transient) + persistent log entry + red status text.
            AppendLog($"\n*** ERROR: {ex.Message} ***");
            ProgressText = "Failed: path too long";
            HasError = true;

            NotificationRequested?.Invoke(
                "Output path too long for Excel",
                ex.Message,
                NotificationKind.Error);

            var killedPath = ExcelProcessTracker.KillAllTracked();
            if (killedPath > 0)
                AppendLog($"\nCleaned up {killedPath} zombie Excel process(es).");
        }
        catch (InsufficientDiskSpaceException ex)
        {
            // Disk-space preflight in BatchEngine.RunAsync. Same UX shape as
            // PathTooLongException — no stack trace clutter, message is already actionable.
            AppendLog($"\n*** ERROR: {ex.Message} ***");
            ProgressText = "Failed: insufficient disk space";
            HasError = true;

            NotificationRequested?.Invoke(
                "Insufficient disk space",
                ex.Message,
                NotificationKind.Error);

            var killedDisk = ExcelProcessTracker.KillAllTracked();
            if (killedDisk > 0)
                AppendLog($"\nCleaned up {killedDisk} zombie Excel process(es).");
        }
        catch (ValidationException vex)
        {
            // Configuration mistake (missing sheet, bad range, corrupt template, etc.).
            // Show a clean user-facing message — no stack trace clutter.
            AppendLog($"\n*** Validation failed ***\n{vex.Message}");
            ProgressText = "Validation failed";
            FlushLog();
            MessageBox.Show(vex.Message, "Calculation Template Validation Failed",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            AppendLog($"\n*** ERROR: {ex.Message} ***");
            AppendLog(ex.StackTrace ?? "");
            ProgressText = "Failed";
            HasError = true;

            NotificationRequested?.Invoke(
                "Batch failed",
                ex.Message,
                NotificationKind.Error);

            // Kill any remaining zombie processes
            int killed = ExcelProcessTracker.KillAllTracked();
            if (killed > 0)
                AppendLog($"\nCleaned up {killed} zombie Excel process(es).");
        }
        finally
        {
            _engine.LogMessage -= OnLogMessage;
            _engine.ProgressChanged -= OnProgressChanged;
            _engine.Dispose();
            _engine = null;
            _logFlushTimer?.Stop();
            _logFlushTimer = null;
            FlushLog(); // Final flush to ensure last messages are shown
            IsRunning = false;
        }
    }

    private bool CanRunBatch() => !IsRunning && !string.IsNullOrWhiteSpace(BatcherFilePath);

    [RelayCommand(CanExecute = nameof(CanCancelBatch))]
    private void CancelBatch()
    {
        // BatchEngine.Cancel already emits "Cancellation requested..." via its own log channel
        // — don't double-log it here.
        _engine?.Cancel();
    }

    private bool CanCancelBatch() => IsRunning;

    private void OnLogMessage(string message)
    {
        // Called from worker threads - just append to buffer, no UI marshalling
        AppendLog(message);
    }

    private void OnProgressChanged(int completed, int total)
    {
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            ProgressPercent = total > 0 ? (double)completed / total * 100 : 0;
            ProgressText = $"{completed} / {total} runs complete ({ProgressPercent:F0}%)";
        });
    }

    /// <summary>
    /// Appends a log message to the buffer (thread-safe, no UI marshalling).
    /// The DispatcherTimer flushes the buffer to the bound LogOutput property at 100ms intervals.
    /// </summary>
    private void AppendLog(string message)
    {
        lock (_logLock)
        {
            _logBuilder.AppendLine(message);

            // Cap in-memory buffer to keep the WPF TextBox responsive on large batches.
            if (_logBuilder.Length > MaxLogChars)
            {
                var trimAt = _logBuilder.Length - (MaxLogChars / 2);
                // Trim on a line boundary near trimAt
                while (trimAt < _logBuilder.Length && _logBuilder[trimAt] != '\n') trimAt++;
                if (trimAt < _logBuilder.Length) trimAt++;
                if (trimAt < _logBuilder.Length)
                {
                    // StringBuilder.Insert(0, ...) shifts the whole buffer (O(n)). Capturing the
                    // tail, clearing, then appending notice+tail does the same number of byte
                    // moves but in one direction — friendlier to the StringBuilder chunk layout.
                    var keep = _logBuilder.ToString(trimAt, _logBuilder.Length - trimAt);
                    _logBuilder.Clear();
                    _logBuilder.Append(LogTruncationNotice);
                    _logBuilder.Append(keep);
                }
                else
                {
                    // The buffer was a single very long line — keep nothing but the notice.
                    _logBuilder.Clear();
                    _logBuilder.Append(LogTruncationNotice);
                }
            }

            _logDirty = true;
        }
    }

    /// <summary>
    /// Flushes buffered log messages to the bound LogOutput property (UI thread only).
    /// </summary>
    private void FlushLog()
    {
        string? snapshot = null;
        lock (_logLock)
        {
            if (_logDirty)
            {
                snapshot = _logBuilder.ToString();
                _logDirty = false;
            }
        }
        if (snapshot != null)
            LogOutput = snapshot;
    }
}







