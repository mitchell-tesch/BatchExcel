using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

namespace BatchExcel.Views;

public partial class AboutDialog
{
    public AboutDialog()
    {
        InitializeComponent();

        // Prefer InformationalVersion so we can show full SemVer (e.g. "0.9.0-beta.1"),
        // not the 4-part AssemblyVersion. Strip the "+<commit>" SourceLink suffix if present.
        var informational = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        string versionText;
        if (!string.IsNullOrWhiteSpace(informational))
        {
            var plus = informational.IndexOf('+');
            versionText = plus >= 0 ? informational[..plus] : informational;
        }
        else
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            versionText = version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "";
        }
        VersionText.Text = string.IsNullOrEmpty(versionText) ? "" : $"Version {versionText}";
    }

    private void OkButton_Click(object sender, RoutedEventArgs e) => Close();

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }
}


