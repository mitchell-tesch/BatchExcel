<div align="center">
  <img src="docs/app_icon.png" alt="BatchExcel" width="150" height="150" />
  <h1>BatchExcel</h1>
  <p><strong>High-performance parallel batch processor for Excel calculation spreadsheets</strong></p>

  ![.NET 10](https://img.shields.io/badge/.NET-10.0-512BD4?style=flat-square&logo=dotnet)
  ![WPF](https://img.shields.io/badge/WPF-Desktop-0078D4?style=flat-square&logo=windows)
  ![Fluent UI](https://img.shields.io/badge/UI-Fluent%20%2F%20Mica-0078D4?style=flat-square&logo=windows11)
  ![Excel COM](https://img.shields.io/badge/Excel-COM%20Interop-217346?style=flat-square&logo=microsoftexcel)
  ![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)

  ---
</div>

Built with C# / WPF / .NET 10, BatchExcel drives multiple Excel instances simultaneously to process large volumes of engineering calculations at maximum speed. Configuration and results live in a simple "batcher" workbook; the heavy lifting is done by parallel COM workers backed by direct OpenXML I/O for everything that doesn't need Excel itself.

## Features

- **Parallel processing** — Distributes batch runs across N Excel instances on dedicated STA threads with a shared `ConcurrentQueue` for automatic load balancing.
- **Worker count is clamped** to `[1, ProcessorCount × 2]` and the number of included runs, so a stray digit can't accidentally spawn dozens of Excel instances.
- **Zombie process protection** — Tracks Excel PIDs via the Win32 `GetWindowThreadProcessId` API and guarantees cleanup on normal exit, cancellation, crashes, or unhandled exceptions.
- **Graceful cancel & shutdown** — Closing the window during a run prompts to cancel, gives workers a short window to exit cleanly, then kills any survivors.
- **COM resilience** — Implements `IOleMessageFilter` for automatic retry on `RPC_E_CALL_REJECTED`, plus per-run retry for transient COM faults; failures on the final attempt are surfaced as **Failed** rows rather than silently dropped.
- **Excel-free I/O** — Reads configuration via ClosedXML and writes results via the OpenXML SDK directly. No Excel process is started for batcher reads/writes, saving ~10 seconds of COM startup/shutdown per batch.
- **Indexed bulk writes** — A custom `SheetWriter` pre-indexes rows and cells so the result-writing pass is O(1) per cell instead of O(rows × cells). Comfortably handles tens of thousands of output cells.
- **User calculation is never modified** — Per-worker copies are made *first*, then the date / job-number headers are written into those copies. The source template on disk is left untouched.
- **Save & PDF export** — Optionally saves a calculated copy of the calculation per run, and/or exports a configured set of sheets to a PDF.
- **Safe CSV output** — Values starting with `=`, `+`, `-`, or `@` are neutralised against CSV/formula injection when opened in Excel/LibreOffice.
- **Modern Fluent UI** — Windows 11 Fluent Design via [WPF-UI](https://github.com/lepoco/wpfui): `FluentWindow` with **Mica** backdrop, Excel-green accent (overrides the system accent), Fluent controls (`ToggleSwitch`, `NumberBox`, `Card`, `ProgressRing`, symbol-icon buttons & menus). The app auto-follows the Windows light/dark theme.
- **Settings persisted** — Last-used batcher path, worker count, save-runs toggle and PDF sheet list are saved to `%AppData%\BatchExcel\settings.json` (debounced so the file isn't rewritten on every keystroke).

## Requirements

- Windows 10 / 11
- .NET 10 SDK (to build) or the .NET 10 Desktop Runtime (to run a published build)
- Microsoft Excel — any version that supports COM automation (Excel 2016+ tested)

## Building

```powershell
dotnet build BatchExcel.sln -c Release
```

To run the tests (xUnit, no Excel required):

```powershell
dotnet test
```

## Usage

1. Launch the application.
2. **File → New Batcher from Template** to create a fresh batcher workbook from the bundled template, or **Browse…** to open an existing one (`.xlsx`).
3. Fill in the batcher workbook in Excel (see [layout](#batcher-workbook-layout) below) and save it.
4. Back in BatchExcel, set:
   - **Parallel Workers** — number of simultaneous Excel instances (typically 2–8 depending on CPU cores and template complexity). Clamped to `[1, ProcessorCount × 2]`.
   - **Save Runs** (optional) — save a calculated copy of the template for each run.
   - **PDF Sheets** (optional) — comma-separated sheet names to include in a per-run PDF export.
5. Click **▶ Start Batch**. Use **■ Cancel** at any time to stop gracefully between runs.

Output is written to a timestamped folder `batch_run_YYMMDD-HHMMSS/` alongside the batcher workbook:

```
batch_run_260529-173057/
├── batch_log.log                                 # Full unbounded log
├── raw_output_fields.csv                         # All results, one row per run
├── <batcher>.xlsx                                # Copy of batcher with results written back
└── <N>_<title>__worker_<W>_<calculation>.xlsx    # Per-run saved copies (if Save Runs enabled)
    <N>_<title>__worker_<W>_<calculation>.pdf     # Per-run PDFs (if PDF Sheets specified)
```

If the original batcher workbook is open in Excel when results are written back, BatchExcel logs a warning and leaves the copy in the output folder as the canonical result.

## Batcher Workbook Layout

The batcher workbook must contain a sheet named **`Main`** with the following layout. Column **A** is unused (label column for the user).

### Calculation Workbook configuration

| Cell | Contents                                                              |
|------|-----------------------------------------------------------------------|
| **B3** | Relative path to the Calculation Spreadsheet (e.g. `Calculator.xlsm`) |
| **B4** | Macro names to execute per run, comma-separated (may be blank)        |
| **B5** | Name of the header sheet in the Calculation Spreadsheet               |

### Project header (written to named ranges in the Calculation Spreadsheet)

| Cell    | Named range written | Description |
|---------|---------------------|-------------|
| **B7**  | `JobNumber`         | Job number |
| **B8**  | `Project`           | Project name |
| **B9**  | `Designer`          | Designer name |
| **B10** | `DesignNotes`       | Design notes |
| _(auto)_ | `Date`             | Batch start time, written automatically |

The Calculation Spreadsheet must define matching workbook-scoped Excel **defined names** for each of these (`Formulas → Define Name…` in Excel). Missing names are silently skipped.

### Batch data table (starts at B15)

| Sheet row | Purpose                                                               |
|-----------|-----------------------------------------------------------------------|
| **15** | Column type: `in` (input), `out` (output), or anything else (skipped) |
| **16** | Calculation Spreadsheet sheet name for that column                    |
| **17** | Cell reference or named range in that sheet                           |
| **18+** | One row per batch run                                                 |

Per-row conventions inside the data table:

- **Column B** — run status. `Yes` to include, anything else (`No`, blank, etc.) to skip.
- **Column C** — run title / identifier (used in CSV output, saved file names, and log messages). Auto-named `Run N` if blank.
- **Column D onwards** — input values for each `in` column, results written into each `out` column.
- An input value of `*` (literal asterisk) is **kept as-is** — the existing template cell value is preserved for that run.

## CSV Output Format

`raw_output_fields.csv` contains every run, with one of three statuses:

| Status      | Meaning |
|-------------|---------|
| `Completed` | Run included and produced results |
| `Skipped`   | Run excluded by `Status ≠ Yes` |
| `Failed`    | Run included but raised an exception (after retries); output columns are blank |

Numbers are written using `InvariantCulture` (decimal point, no thousands separator). Strings beginning with `=`, `+`, `-`, or `@` are prefixed with a single quote so they aren't interpreted as formulas when the CSV is opened in Excel/LibreOffice.

## Architecture

```
BatchExcel/
├── Models/
│   └── BatchConfig.cs               # Data models (BatchConfig, FieldDefinition, BatchRun)
├── Services/
│   ├── BatchEngine.cs               # Parallel batch orchestrator + log file management
│   ├── ExcelWorker.cs               # Per-worker COM loop (cached refs, retry, RCW release)
│   ├── ComMessageFilter.cs          # IOleMessageFilter for retry on COM busy
│   ├── ExcelProcessTracker.cs       # PID tracking + safe quit + zombie cleanup
│   ├── BatcherReader.cs             # Read config (ClosedXML), write results (OpenXML)
│   ├── CalculationHeaderWriter.cs   # Write date/job/etc. headers via OpenXML defined names
│   ├── CalculationValidator.cs      # Dry-run validation of calculation template (sheets + ranges)
│   ├── OpenXmlHelpers.cs            # SheetWriter (indexed bulk writes) + cell helpers
│   ├── CsvResultWriter.cs           # CSV output with injection protection
│   ├── PdfExporter.cs               # Multi-sheet PDF export via Excel COM
│   ├── FileNameSanitizer.cs         # Invalid-char replacement for output file names
│   └── UserSettings.cs              # JSON-persisted user preferences
├── ViewModels/
│   └── MainViewModel.cs             # MVVM ViewModel (CommunityToolkit.Mvvm)
├── Views/
│   └── AboutDialog.xaml(.cs)        # Help → About dialog (Fluent FluentWindow)
├── Themes/
│   └── ModernTheme.xaml             # Fonts + Excel-green accent palette overriding WPF-UI defaults
├── MainWindow.xaml(.cs)             # FluentWindow + Mica + auto-scroll + graceful close
├── App.xaml(.cs)                    # WPF-UI themes merged + system-theme follower + zombie cleanup
└── Resources/                       # App icon (multi-frame ICO + 2048 PNG) + bundled batcher template
```

### Processing pipeline

```
1. Read config         BatcherReader.ReadConfig          ClosedXML, direct file (~50 ms)
2. Validate template   CalculationValidator.Validate     OpenXML, fail-fast dry run (~10 ms)
3. Copy Calculation    one .xlsx per worker              File.Copy × N
4. Write headers       CalculationHeaderWriter.Write     OpenXML on each copy (~20 ms)
5. Launch workers      RunOnStaThread                    STA + IOleMessageFilter, staggered
6. Process runs        ExcelWorker.ProcessRunQueue       Cached COM refs → Calculate() → read
                                                         + per-run Save & PDF (optional)
7. Release & quit      ExcelWorker finally               FinalReleaseComObject + SafeQuitExcel
8. Write results       BatcherReader.WriteResults        Single OpenXML pass → File.Copy ×2
9. Write CSV           CsvResultWriter.Write             StreamWriter, injection-safe escaping
10. Cleanup            Delete worker copies              KillAllTracked on exit / on errors
```

### Performance design

| Technique                                                                    | Effect |
|------------------------------------------------------------------------------|--------|
| Per-worker calculation file copy                                             | Eliminates file locking contention between Excel instances |
| Cached sheet + range COM references                                          | Eliminates repeated COM lookups per run |
| Explicit `Marshal.FinalReleaseComObject` on cached refs                      | Reliable Excel shutdown without leaking RCWs |
| Manual calc mode + `Calculate()` per run                                     | Only recalculates dirty cells |
| Disabled: Events, ScreenUpdating, AutoRecover, Interactive, AskToUpdateLinks | Removes background overhead |
| `IOleMessageFilter` + per-run retry on transient `COMException`s             | Automatic recovery from Excel-busy errors |
| OpenXML for batcher I/O                                                      | Eliminates 2 Excel process lifecycles per batch |
| Indexed `SheetWriter` for result writes                                      | O(1) per cell instead of O(rows × cells) |
| Single OpenXML pass + `File.Copy` mirror                                     | Result writes once, mirrored to original + output folder |
| Staggered worker startup (200 ms × workerId)                                 | Smooths out Excel process launch CPU spike |
| Throttled progress reporting (100 ms)                                        | Prevents UI dispatcher flooding |
| In-memory log cap (256 KB)                                                   | Keeps WPF `TextBox` responsive on long batches (full log on disk) |
| `ConcurrentQueue` work distribution                                          | Self-balancing across workers |

## Testing

The `BatchExcel.Tests` project (xUnit) currently includes **81 tests** covering:

- `BatcherReader` — round-trips against generated `.xlsx` fixtures (no Excel required)
- `OpenXmlHelpers` / `SheetWriter` — cell reference math, indexed bulk writes, typed-value coverage (`double` / `int` / `long` / `float` / `decimal` / `bool` / `DateTime`)
- `CsvResultWriter` — escaping, status distinction, formula-injection neutralisation, invariant-culture numbers
- `BatchConfig` — macro parsing, included-run count
- `CalculationValidator` — dry-run validation: sheet existence, A1 and named-range resolution, error aggregation, corrupt-file handling
- `FileNameSanitizer` — invalid character handling
- `UserSettings` — load / save round-trip

```powershell
dotnet test
```

## Dependencies

- [CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet) 8.4 — MVVM source generators (`ObservableObject`, `[RelayCommand]`)
- [WPF-UI](https://github.com/lepoco/wpfui) 4.3 — Fluent Design controls (`FluentWindow`, `Card`, `Button`, `TextBox`, `NumberBox`, `ToggleSwitch`, `ProgressRing`, `HyperlinkButton`, `SymbolIcon`, `MessageBox`) and the Mica backdrop / system-theme follower
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) 0.105 — Convenient Excel reading API over OpenXML
- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK) 3.1 — Direct Excel file writing (pinned explicitly so a future ClosedXML drop can't silently change the version we write against)

## License

This project is licensed under the [MIT License](LICENSE).

