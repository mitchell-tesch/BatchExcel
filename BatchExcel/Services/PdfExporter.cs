namespace BatchExcel.Services;

using System.IO;
using System.Runtime.InteropServices;

/// <summary>
/// Exports specified sheets of a spreadsheet to PDF via Excel COM interop.
/// Disables printer driver communication during the export for major speedup.
/// </summary>
internal static class PdfExporter
{
    /// <summary>
    /// Exports the specified sheets of the workbook to a single PDF file.
    /// PDF export failures do not abort the surrounding batch run — they are reported via
    /// <paramref name="log"/> so the user has a record of which exports failed and why.
    /// </summary>
    public static void Export(dynamic workbook, List<string> sheetNames, string pdfPath, Action<string>? log = null)
    {
        var sheets = new List<dynamic>();
        dynamic? app = null;
        var printCommToggled = false;

        try
        {
            // NOTE: workbook.Application returns the same underlying COM object as the
            // ExcelWorker's cached excelApp RCW. We must NOT FinalReleaseComObject it here
            // — doing so detaches the worker's RCW and the next call on excelApp throws
            // "COM object that has been separated from its underlying RCW cannot be used."
            app = workbook.Application;

            // Resolve sheet references, recording any names that couldn't be found so the user
            // can spot a typo in the PDF Sheets setting.
            var missing = new List<string>();
            foreach (var name in sheetNames)
            {
                try { sheets.Add(workbook.Sheets[name]); }
                catch { missing.Add(name); }
            }
            if (missing.Count > 0)
                log?.Invoke($"\tPDF: skipped missing sheet(s): {string.Join(", ", missing)}");

            if (sheets.Count == 0)
            {
                log?.Invoke($"\tPDF: no matching sheets found for '{Path.GetFileName(pdfPath)}', export skipped.");
                return;
            }

            // Select all target sheets. Activate the first so ExportAsFixedFormat targets it,
            // then additively select the rest.
            sheets[0].Select();
            for (int i = 1; i < sheets.Count; i++)
            {
                sheets[i].Select(Replace: false);
            }

            // Disable printer driver communication during the slow export call for major speedup
            try { app.PrintCommunication = false; printCommToggled = true; }
            catch
            {
                // ignored — older Excel versions / non-standard configurations may not expose it
            }

            workbook.ActiveSheet.ExportAsFixedFormat(
                Type: 0, // xlTypePDF
                Filename: pdfPath,
                Quality: 0, // xlQualityStandard
                IncludeDocProperties: false,
                IgnorePrintAreas: false);
        }
        catch (Exception ex)
        {
            // PDF export failure should not stop the batch, but it must not be silent either —
            // a missing PDF with no log entry would be very confusing for the user.
            log?.Invoke($"\tPDF export failed for '{Path.GetFileName(pdfPath)}': {ex.Message}");
        }
        finally
        {
            if (printCommToggled)
            {
                try { app!.PrintCommunication = true; }
                catch { /* ignored */ }
            }

            // Restore single-sheet selection so subsequent runs aren't operating against
            // a multi-sheet group (which can change the semantics of some COM writes).
            if (sheets.Count > 0)
            {
                try { sheets[0].Select(); } catch { /* ignored */ }
            }
        }
    }
}

