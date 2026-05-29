using System.ComponentModel;
using System.Windows;
using BatchExcel.Services;
using BatchExcel.ViewModels;
using BatchExcel.Views;

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
        }
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
                PrimaryButtonAppearance = Wpf.Ui.Controls.ControlAppearance.Danger,
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