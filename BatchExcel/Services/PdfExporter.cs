namespace BatchExcel.Services;

using System.Runtime.InteropServices;

/// <summary>
/// Exports specified sheets of a spreadsheet to PDF via Excel COM interop.
/// Disables printer driver communication during the export for major speedup.
/// </summary>
internal static class PdfExporter
{
    /// <summary>
    /// Exports the specified sheets of the workbook to a single PDF file.
    /// </summary>
    public static void Export(dynamic workbook, List<string> sheetNames, string pdfPath)
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

            // Resolve sheet references, silently skipping any that don't exist
            foreach (var name in sheetNames)
            {
                try { sheets.Add(workbook.Sheets[name]); } catch { /* skip missing */ }
            }

            if (sheets.Count == 0) return;

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
                // ignored
            }

            workbook.ActiveSheet.ExportAsFixedFormat(
                Type: 0, // xlTypePDF
                Filename: pdfPath,
                Quality: 0, // xlQualityStandard
                IncludeDocProperties: false,
                IgnorePrintAreas: false);
        }
        catch
        {
            // PDF export failure should not stop the batch
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

