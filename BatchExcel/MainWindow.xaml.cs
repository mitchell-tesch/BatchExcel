using System.ComponentModel;
using System.Windows;
using BatchExcel.Services;
using BatchExcel.ViewModels;
using BatchExcel.Views;
using Wpf.Ui.Controls;

namespace BatchExcel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
    public MainWindow()
    {
        InitializeComponent();

        // Auto-scroll log to bottom when text changes
        if (DataContext is MainViewModel vm)
        {
            vm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(MainViewModel.LogOutput))
                {
                    LogTextBox.ScrollToEnd();
                }
            };

            // Forward VM notifications to the Fluent snackbar overlay.
            vm.NotificationRequested += OnNotificationRequested;
        }
    }

    private void OnNotificationRequested(string title, string message, MainViewModel.NotificationKind kind)
    {
        // Marshal to UI thread — VM raises this from the awaited Task continuation, which is
        // usually the UI thread, but guard anyway in case a future code path raises off-thread.
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(() => OnNotificationRequested(title, message, kind));
            return;
        }

        var (appearance, icon, timeoutSeconds) = kind switch
        {
            MainViewModel.NotificationKind.Error =>
                (ControlAppearance.Danger,  SymbolRegular.ErrorCircle24, 10),
            MainViewModel.NotificationKind.Warning =>
                (ControlAppearance.Caution, SymbolRegular.Warning24,     7),
            _ =>
                (ControlAppearance.Success, SymbolRegular.Checkmark24,   4),
        };

        var snackbar = new Snackbar(RootSnackbar)
        {
            Title = title,
            Content = message,
            Appearance = appearance,
            Icon = new SymbolIcon(icon),
            Timeout = TimeSpan.FromSeconds(timeoutSeconds),
        };
        snackbar.Show();
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        // If a batch is running, ask the user and request a graceful cancel before killing processes.
        if (DataContext is MainViewModel vm && vm.IsRunning)
        {
            var msg = new Wpf.Ui.Controls.MessageBox
            {
                Title = "BatchExcel",
                Content = "A batch is still running. Cancel it and close?",
                PrimaryButtonText = "Cancel & close",
                PrimaryButtonAppearance = ControlAppearance.Danger,
                CloseButtonText = "Keep running",
            };

            // ShowDialog is sync; ShowDialogAsync would race the e.Cancel write.
            var result = msg.ShowDialogAsync().GetAwaiter().GetResult();
            if (result != Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                e.Cancel = true;
                return;
            }

            // Request cancellation and give workers a short window to exit cleanly.
            vm.CancelBatchCommand.Execute(null);

            // Pump the dispatcher briefly so async continuations can run and workers can stop.
            var deadline = DateTime.UtcNow.AddSeconds(3);
            while (vm.IsRunning && DateTime.UtcNow < deadline)
            {
                Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                Thread.Sleep(50);
            }
        }

        // Kill any zombie Excel processes on window close
        int killed = ExcelProcessTracker.KillAllTracked();
        if (killed > 0)
        {
            var info = new Wpf.Ui.Controls.MessageBox
            {
                Title = "BatchExcel Cleanup",
                Content = $"Cleaned up {killed} Excel process(es) that were still running.",
                CloseButtonText = "OK",
            };
            _ = info.ShowDialogAsync();
        }
    }

    private void Exit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void About_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new AboutDialog { Owner = this };
        dialog.ShowDialog();
    }
}