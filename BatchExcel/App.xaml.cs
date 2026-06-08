using System.Windows;
using System.Windows.Media;
using BatchExcel.Services;
using Wpf.Ui.Appearance;

namespace BatchExcel;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App
{
    // Excel-green brand accent. Overrides the Windows system accent everywhere
    // WPF-UI consults Application accent resources (Primary buttons, focus
    // rings, sliders, toggle switches, progress bars, hyperlinks, ...).
    private static readonly Color BrandAccent = Color.FromRgb(0x1F, 0x7A, 0x47);

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Follow the Windows light/dark system theme automatically, BUT do not
        // let WPF-UI grab the system accent — we want to stay Excel-green.
        ApplicationThemeManager.ApplySystemTheme(updateAccent: false);
        ApplyBrandAccent();

        // Re-apply our accent whenever WPF-UI swaps the theme at runtime
        // (e.g. user toggles Windows dark mode), since theme changes can
        // refresh the Application accent resource dictionary.
        ApplicationThemeManager.Changed += (_, _) => ApplyBrandAccent();

        // SystemThemeWatcher.Watch needs the MainWindow's HWND, but Current.MainWindow is
        // still null here because StartupUri navigation happens AFTER OnStartup returns.
        // Defer the watch hook-up until the main window actually exists.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (Current.MainWindow != null)
            {
                SystemThemeWatcher.Watch(Current.MainWindow, updateAccents: false);
            }
        }), System.Windows.Threading.DispatcherPriority.Loaded);

        // Global exception handlers to ensure zombie Excel cleanup
        AppDomain.CurrentDomain.UnhandledException += (_, _) =>
        {
            ExcelProcessTracker.KillAllTracked();
        };

        DispatcherUnhandledException += (_, args) =>
        {
            ExcelProcessTracker.KillAllTracked();
            var msg = new Wpf.Ui.Controls.MessageBox
            {
                Title = "BatchExcel Error",
                Content = $"An unexpected error occurred:\n\n{args.Exception.Message}",
                CloseButtonText = "OK",
            };
            msg.ShowDialogAsync();
            args.Handled = true;
        };
    }

    private static void ApplyBrandAccent()
    {
        // WPF-UI generates Primary / Secondary / Tertiary shades from the seed
        // colour and pushes them into the Application accent resource keys.
        ApplicationAccentColorManager.Apply(
            BrandAccent,
            ApplicationThemeManager.GetAppTheme());
    }

    protected override void OnExit(ExitEventArgs e)
    {
        ExcelProcessTracker.KillAllTracked();
        base.OnExit(e);
    }
}



