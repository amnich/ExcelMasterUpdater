# Excel Master Updater

A high-performance, zero-dependency PowerShell and .NET application designed to maintain and update a single **Master ("Base") Excel file** of children/student records from incoming change files.

Built as an architectural sibling of `Compare-ExcelFiles.ps1`, it introduces row-by-row change review, header fingerprint auto-detection, InPlace OpenXML stream patching, metadata stamping, audit logging, automated backups, and resilient CSV handling.

---

## Key Features

1. **Header Fingerprint Profile Store & Real-Time Search**:
   - Computes SHA-256 header fingerprints from incoming worksheets (`HeaderFingerprint = sha256(join("|", normalized headers))`).
   - Automatically recognizes previously mapped source templates and loads mappings silently.
   - Supports 1:N delimiter/regex splits and N:1 column merges.
   - **Mapping Rules Search & Filter**: Real-time filtering (`txtSearchMapping`) by base or incoming column names with instant clear (`✖`).
   - **Reactive Live Preview & Join Key Indicator**: Instant sample re-projection when changing join keys, with key icon (**`🔑 `**) in the Target Base Column.
2. **Review & Approval Engine (Row-by-Row)**:
   - Categorizes incoming rows into `New`, `Changed`, `Unchanged`, `Ambiguous`, and `Removed`.
   - Side-by-side / stacked diff inspector with per-cell modified highlights.
   - Selective cell updates: users can deselect specific cells they do not wish to update.
   - **Double-Click Quick Edit**: Double-clicking any record in the review items list directly opens the field override dialog.
   - Ambiguous resolution for duplicate keys, missing keys, or multi-row candidates.
   - Full keyboard navigation: `[A]` Accept, `[R]` Reject, `[S]` Skip, `[E]` Edit Value override, `[B]` / `[←]` Back navigation, `[Ctrl+Z]` Undo last decision, with auto-scrolling (`ScrollIntoView`).
   - Bulk batch actions: `Accept All` (safely skips ambiguous records) and `Reject All`.
   - Export Review Report: Exports complete discrepancy and review summaries directly to formatted `.xlsx`, responsive `.html`, or delimited `.csv`.
3. **InPlace OpenXML Engine & Resilient File Ingestion**:
   - Patches only modified cells using OpenXML `inlineStr` without touching `sharedStrings.xml`.
   - Completely preserves worksheet formatting, cell styles, fonts, colors, and extra sheets.
   - Appends new accepted records to `<sheetData>` with running row numbers.
   - Atomic replacement via temporary files (`<file>.tmp.xlsx`).
   - Concurrency & crash safety: aborts cleanly if base file is locked; auto-detects orphaned `.tmp.xlsx` files.
   - Built-in structure guards: detects unsupported `<tableParts>`, `<pivotTable>`, or shared formulas and **automatically falls back to SafeRewrite mode** with a status-bar advisory.
   - **Resilient CSV Engine**: Automatically detects delimiters (semicolon `;`, comma `,`, tab `\t`, pipe `|`) and character encodings (UTF-8 with/without BOM, UTF-16, and Windows-1250 Central European ANSI fallback).
4. **Metadata Stamping Engine & Tabbed Settings**:
   - Stamped on every accepted change or added row:
     - `ChangeDate`: timestamp (e.g. `yyyy-MM-dd HH:mm`).
     - `SourceFileFullPath` & `SourceFileName`.
     - `CurrentUser`: `$env:USERNAME`.
     - `SourceRowNumber`: physical row number in incoming source.
     - `ImportBatchId`: session GUID.
   - Missing metadata columns are automatically created in header row 1 on first write.
   - Interactive Settings Dialog with tabs for General preferences and a live DataGrid editor for Metadata columns and tokens.
5. **Audit Logging & Automatic Backup**:
   - Detailed per-session logs in `JSONL` and human-readable `TXT` format.
   - Append-only `masterlog.jsonl` audit trail.
   - Optional change audit worksheet (`ImportLog`) appended directly inside the master Excel workbook.
   - Automatic pre-write backups in `Backups/` with configurable retention count (default 20) and **auto-prune by total folder size** (default 200 MB threshold, configurable via `MaxBackupMb`).
   - Backup size warning surfaced in the success dialog if auto-pruning occurred.
6. **Headless Batch Pipeline & Reporting**:
   - `-Headless` CLI pipeline mode for automated unattended imports.
   - Configurable `-AutoAccept` policies: `AllNonAmbiguous`, `OnlyChanged`, `OnlyNew`, or `None`.
   - Standalone responsive HTML Diff Report generator (`-ExportReportPath "report.html"`) with styled KPI cards and discrepancy tables.
   - **Machine-Readable Execution Summary**: Automatically outputs `last_execution_summary.json` (or custom `-SummaryJsonPath`) detailing counts of new, changed, ambiguous, unchanged, accepted rows, updated cells, and backup paths.
7. **Folder-Watch Daemon Mode**:
   - `Watch-IncomingFolder.ps1`: monitors an incoming drop folder using file-readiness locking checks.
   - Automatically processes incoming `.xlsx` / `.csv` files and routes finished files to `Archive/` or `Failed/`.
   - One-shot mode (`-Once`) for scheduled tasks and automated CI/batch runs.
8. **Zero External Dependencies & Standalone Binary**:
   - Requires no Excel COM, no Microsoft Office installed, and no `ImportExcel` module.
   - Built with pure .NET `System.IO.Compression` and `System.Xml`.
   - Fully compatible with Windows PowerShell 5.1 (.NET Framework) and PowerShell 7+ (.NET Core).
   - Compilable to standalone windowed EXE via `ps2exe` with embedded language catalog and icon.
9. **Trilingual Dynamic Localization**:
   - 283 localization keys maintained across Polish (`pl`), English (`en`), and German (`de`) with 100% key parity.
   - Dynamic real-time language switching without application restart.

---

## Security & PII Notice

> [!WARNING]
> **Data Privacy (GDPR / RODO)**:
> This application processes sensitive personal identification data (children and parent names, PESEL numbers, contact numbers, addresses, and dietary needs).
> - Ensure the `LogDirectory` (`%APPDATA%\MasterUpdater\Logs` by default) and `BackupDirectory` are placed on restricted storage volumes protected by strict NTFS Access Control Lists (ACLs).
> - To mask names and addresses in generated logs, enable the `RedactNamesInLog` option in Settings or `config.json` (`"RedactNamesInLog": true`).

---

## Directory Layout

```
UpdateExcelBaseFileProject/
├── start.ps1                       # Entry-point launcher (STA thread apartment guard, -UseExe)
├── Master-Updater.ps1              # Main single-script module (GUI & Headless engine)
├── Watch-IncomingFolder.ps1        # Folder-watch daemon & auto-import runner
├── Master-Updater.exe              # Standalone compiled executable (PS2EXE)
├── Build-Exe.ps1                   # Standalone PS2EXE build script
├── language.json                   # Trilingual localization catalog (283 keys: PL / EN / DE)
├── README.md                       # Main architecture & user manual
├── Docs/                           # Comprehensive documentation
│   ├── USER_GUIDE.md               # Trilingual unified user guide (EN / PL / DE)
│   ├── USER_GUIDE_EN.md            # English user guide
│   ├── USER_GUIDE_PL.md            # Polish user guide (Instrukcja obsługi)
│   ├── USER_GUIDE_DE.md            # German user guide (Benutzerhandbuch)
│   └── ARCHITECTURE.md             # Technical architecture specification
├── HelperDocs/                     # Strategic reviews
│   └── swot_analysis.md            # Project SWOT review & priority matrix
├── TestFixtures/                   # Fixture generation & test files
│   ├── Generate-TestFixtures.ps1   # Fixture builder
│   ├── base.xlsx                   # 50 rows master base file
│   ├── base_formatted.xlsx         # Multi-sheet styled base file
│   ├── incoming_A.xlsx             # 50 rows incoming (3 changed, 2 new, 45 unchanged)
│   └── incoming_B.csv              # Edge cases (empty join key, duplicate incoming key)
└── Tests/                          # Automated test suites (Dual PS 5.1 & PS 7)
    ├── Test-BaseSheetLogging.ps1           # Audit worksheet writing and column sanity
    ├── Test-CompareEngine.ps1              # Comparison, projection, ambiguous & removed detection
    ├── Test-DataMappingPreview.ps1         # Live sample projection and status badges
    ├── Test-DeepVerification.ps1           # Deep integrity and edge case checks
    ├── Test-E2E-MasterUpdater.ps1          # Complete end-to-end integration test
    ├── Test-EditExcelHelper.ps1            # InPlace XML patching & row appending
    ├── Test-GridSplitterBehavior.ps1       # UI splitter layout and persistence
    ├── Test-HeadlessAndWatcher.ps1         # Headless batch, HTML reporting, and folder watcher
    ├── Test-Localization.ps1               # 100% key completeness across PL / EN / DE (283 keys)
    ├── Test-NegativePath.ps1               # Fuzz & negative-path tests (corrupt ZIP, locks)
    ├── Test-NewUXAndResilienceFeatures.ps1 # Tab 1 search, double-click, CSV pipe & Win-1250
    ├── Test-SampleFilesIntegration.ps1     # Multi-format integration tests
    ├── Test-UIComponents.ps1               # Automated XAML UI & non-interactive launch
    └── Test-UnchangedOption.ps1            # Show unchanged toggle and filtering
```

---

## Usage

### Quick Start (GUI)
To launch the graphical interface:
```powershell
pwsh -File .\start.ps1
```
Or via Windows PowerShell 5.1:
```powershell
powershell -File .\start.ps1
```
Or launch the standalone binary directly:
```powershell
.\Master-Updater.exe
```

### Headless Batch Import (Unattended Pipeline)
Run automated comparison and write-back without displaying the GUI:
```powershell
pwsh -File .\Master-Updater.ps1 -BaseFilePath "base.xlsx" -IncomingPath "incoming.csv" -Headless -AutoAccept AllNonAmbiguous -ExportReportPath "diff_report.html" -SummaryJsonPath "summary.json"
```

### Folder Watcher Daemon
Monitor an incoming folder continuously:
```powershell
pwsh -File .\Watch-IncomingFolder.ps1 -BaseFilePath "base.xlsx" -FolderPath ".\IncomingDrop" -ExportHtmlReports
```
Or run a one-shot drain:
```powershell
pwsh -File .\Watch-IncomingFolder.ps1 -BaseFilePath "base.xlsx" -FolderPath ".\IncomingDrop" -Once
```

---

## Building Standalone Executable
To compile `Master-Updater.ps1` into a standalone Windows executable (`Master-Updater.exe`) with embedded `language.json`:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Build-Exe.ps1
```

---

## Automated Verification Suite
To execute the automated regression test matrix:
```powershell
pwsh -File .\Tests\Test-Localization.ps1
pwsh -File .\Tests\Test-NewUXAndResilienceFeatures.ps1
pwsh -File .\Tests\Test-EditExcelHelper.ps1
pwsh -File .\Tests\Test-CompareEngine.ps1
pwsh -File .\Tests\Test-E2E-MasterUpdater.ps1
pwsh -File .\Tests\Test-HeadlessAndWatcher.ps1
```
All 14 test suites are 100% compatible and validated on both Windows PowerShell 5.1 and PowerShell 7.
