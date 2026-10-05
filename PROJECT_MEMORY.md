# Excel Master Updater — Project Memory & Technical Blueprint

> **File:** `\ExcelMasterUpdater\PROJECT_MEMORY.md`  
> **Status:** Production Reference / System Living Specification  
> **Encoding:** UTF-8 with BOM (`0xEF, 0xBB, 0xBF`)  
> **Compatibility:** Windows PowerShell 5.1 & PowerShell 7.x (Core)  
> **Deployment Target:** Standalone GUI Executable (`Master-Updater.exe` via `PS2EXE`) & Headless Automation  

---

## 1. Executive Summary & Purpose

**Excel Master Updater** is a mission-critical, enterprise-grade data reconciliation and synchronization platform built for Windows. It enables operators to maintain a definitive **Master ("Base") Excel file** (e.g., student rosters, employee directories, product registries) updated from recurring incoming update files (e.g., monthly admissions, CRM dumps, contractor exports).

### Core Problem Solved
Traditional Excel reconciliation suffers from:
1. **Destructive Overwrites**: Standard Excel libraries (or manual copy-pasting) strip formatting, macros, data validation, and additional sheets.
2. **Heavy External Dependencies**: Reliance on Microsoft Office Excel COM automation (`Excel.Application`) or third-party modules (`ImportExcel`, `EPPlus`) creates deployment friction, licensing overhead, and COM stability leaks.
3. **Black-Box Errors**: Operators cannot preview row-by-row differences, selectively apply field changes, or maintain an indisputable audit trail.

### Solution Architecture
Excel Master Updater provides:
- **Zero External Dependencies**: Powered purely by native .NET Framework / .NET Core runtime assemblies (`System.IO.Compression`, `System.Xml`, `System.Xml.Linq`, WPF).
- **InPlace OpenXML Stream Patching**: Directly mutates cell nodes in the OpenXML ZIP package without touching cell styles, colors, fonts, formulas, column widths, or companion worksheets.
- **Fingerprint Profile Mapping**: Computes SHA-256 header fingerprints to instantly auto-load mapping rules for recurring monthly layouts.
- **Trilingual GUI**: 100% localized interface (English, Polish, German) with real-time on-the-fly language switching and dark/light themes.
- **Fail-Safe Write Engine**: Mandatory pre-write backups, atomic `.tmp.xlsx` staging with schema verification before disk commit, crash recovery hooks, and optional in-workbook audit sheet logging (`ImportLog`).

---

## 2. Technical Stack & Runtime Requirements

| Layer | Component / Standard | Notes |
|---|---|---|
| **OS** | Windows 10 / 11 / Windows Server 2016+ | 64-bit architecture recommended. |
| **PowerShell Engines** | Windows PowerShell 5.1 (.NET Framework 4.7.2+)<br>PowerShell 7.x (.NET Core 6.0+) | 100% dual-engine runtime compatibility verified across all features and tests. |
| **GUI Framework** | Windows Presentation Foundation (WPF) | Embedded XAML with dynamic runtime binding and thread-safe UI updates. |
| **File Formats** | Excel Workbook (`.xlsx`), CSV (`.csv`) | OpenXML ZIP reader/patcher for `.xlsx`; high-performance regex RFC-4180 parser for `.csv`. |
| **Packaging** | `PS2EXE` (v1.0.18+) | Compiles `Master-Updater.ps1` to windowed `Master-Updater.exe` with embedded `language.json`. |
| **Character Encoding** | UTF-8 with BOM (`0xEF, 0xBB, 0xBF`) | Mandatory across all script files, test fixtures, documentation, and config catalogs. |

---

## 3. Repository & Directory Structure

```
\ExcelMasterUpdater\
├── Master-Updater.ps1        # Monolithic production script (Regions 1-10, ~9,100 lines)
├── Master-Updater.exe        # Compiled standalone binary (PS2EXE, 771 KB)
├── start.ps1                 # Universal launcher (dual PS5.1/PS7 engine & parameter forwarder)
├── Build-Exe.ps1             # Automated compilation script utilizing PS2EXE
├── language.json             # Trilingual catalog (EN, PL, DE — 323 keys per language)
├── config.json               # Application runtime preferences & persistence
├── PROJECT_MEMORY.md         # This technical specification and memory file
├── Docs\                     # Documentation suite
│   ├── ARCHITECTURE.md       # High-level architectural specification
│   ├── USER_GUIDE.md         # Composite trilingual manual
│   ├── USER_GUIDE_EN.md      # Standalone English user guide
│   ├── USER_GUIDE_PL.md      # Standalone Polish user guide
│   └── USER_GUIDE_DE.md      # Standalone German user guide
├── Backups\                  # Automated pre-write backup store (<BaseName>.<Timestamp>.bak.xlsx)
├── Logs\                     # JSONL and formatted text execution audit trails
├── Profiles\                 # Saved column mapping profiles (JSON v2.0 format)
└── Tests\                    # Comprehensive automated test suite (18 test files)
    ├── Test-Localization.ps1              # 100% trilingual key completeness (320 keys) & Get-UiString coverage (432 calls)
    ├── Test-QoLSelectiveReconciliation.ps1 # Diff policies, selective field takeover, inline editing, revert & new file write
    ├── Test-RealUserFilesDDOIntegration.ps1# Real-world student admission file reconciliation (DDO 2 dataset)
    ├── Test-NewUXAndResilienceFeatures.ps1 # Mapping search filter, ANSI/CP1250 encoding, join key badges
    ├── Test-DataMappingPreview.ps1         # Live mapping results preview, sample row grids, row navigation
    ├── Test-GridSplitterBehavior.ps1       # Interactive UI resizing and splitter stability
    ├── Test-ConcatenateMappingRule.ps1     # Multi-column concatenation & delimiter handling
    ├── Test-CompareOptions.ps1             # Case sensitivity, whitespace, and diacritics options
    ├── Test-CompareEngine.ps1              # Join keys, composite keys, status classifications
    ├── Test-EditExcelHelper.ps1            # OpenXML InPlace patching & row appending
    ├── Test-HeadlessAndWatcher.ps1         # CLI automation, non-interactive batch runs
    ├── Test-NegativePath.ps1               # Corrupted files, missing sheets, invalid keys
    ├── Test-SampleFilesIntegration.ps1     # Real-world fixtures (Students, Grades, Attendance)
    ├── Test-UIComponents.ps1               # WPF XAML element tree & control instantiation
    ├── Test-UnchangedOption.ps1            # Show/hide unchanged records & review filters
    ├── Test-BaseSheetLogging.ps1           # In-workbook ImportLog audit sheet generation & PII
    ├── Test-E2E-MasterUpdater.ps1          # End-to-end full reconciliation and write-back cycle
    └── Test-DeepVerification.ps1           # Extreme multi-sheet, formula, and UTF-8 verification
```

---

## 4. Code Architecture: Regions of `Master-Updater.ps1`

`Master-Updater.ps1` is organized into 10 structured regions:

### Region 1: Dependencies & Embedded C# Classes
Compiles low-overhead C# helper classes directly into the runtime memory via `Add-Type`:
- **`FastExcelHelper`**: Reads `.xlsx` worksheets by streaming XML nodes (`XmlReader`) from `xl/worksheets/sheetN.xml` and unzipping `xl/sharedStrings.xml`. Implements fast CSV parsing with RFC-4180 quote escaping. Exports datasets to clean `.xlsx` packages without COM dependencies.
- **`FastDiffHelper`**: High-performance string normalization engine. Provides methods for whitespace collapsing, diacritic/special character stripping, and exact case-sensitive/case-insensitive comparisons.
- **`EditExcelHelper`**: High-speed, zero-dependency OpenXML patcher. Operates directly on the `.xlsx` ZIP container:
  - Updates cell contents using inline strings (`<c r="C5" t="inlineStr"><is><t>val</t></is></c>`), preserving cell style attributes (`s`), colors, fonts, and borders.
  - Appends new rows sequentially into `<sheetData>`, inherits column style attributes (`s`) from previous rows, and updates worksheet `<dimension>` bounds.
  - Automatically discovers and expands Excel Table boundaries (`ref` and `autoFilter.ref` in `xl/tables/table*.xml`) to encompass newly appended rows, preserving table styles, alternating row stripes, and filters.
  - Guaranteed 100% InPlace patching without destructive fallback to `SafeRewrite` for `.xlsx` files.

### Region 2: Configuration Management
- Manages application settings stored in `config.json` (or `%APPDATA%\MasterUpdater\config.json`).
- Handled by `Get-AppConfig` and `Save-AppConfig`.
- Configurable properties include:
  - `DefaultLanguage` (`en`, `pl`, `de`)
  - `Theme` (`Dark`, `Light`)
  - `BackupDirectory` & `BackupRetentionCount` (default: 20)
  - `MaxBackupMb` (default: 200 MB disk limit)
  - `WriteMode` (`InPlace` vs `SafeRewrite`)
  - `RedactNamesInLog` (PII anonymization toggle)
  - `DetectRemovedRows` (detect missing records)
  - `ShowUnchangedRows` (review list visibility for identical records)
  - `LogToBaseSheet` (append audit trail to workbook sheet)
  - `BaseLogSheetName` (default: `ImportLog`)

### Region 3: Profile Store & Fingerprints
- Computes SHA-256 header signatures via `Compute-HeaderFingerprint`:
  $$\text{Fingerprint} = \text{SHA256}(\text{SheetName} + \text{"\#"} + \text{Join}("|", \text{NormalizedHeaders}))$$
- Handled by `Get-MappingProfiles`, `Save-MappingProfile`, and `Find-MappingProfileByFingerprint`.
- Enables automatic template detection when loading recurring spreadsheets.

### Region 4: Mapping Projection Engine
- Handled by `Get-ProjectedRow` and `Build-JoinKey`.
- Supports complex schema alignments:
  - **1:1 Mapping**: Direct column-to-column passthrough.
  - **N:1 Merge**: Combines multiple source fields into a single base column using `Exact`, `FirstNonEmpty`, or `Concatenate` modes (with custom separators).
  - **1:N Split**: Splits a compound source column (e.g., `City, Street 10`) across multiple base fields via delimiter or regex.
  - **Composite Keys**: Builds normalized compound lookup keys (e.g., `[Smith]|[John]|[1995-05-12]`).

### Region 5: Compare Engine
- Handled by `Invoke-MasterCompare`.
- Processes the base file and incoming file in memory, yielding a collection of `ReviewItem` objects classified as:
  - **`New`**: Incoming join key does not exist in the master base. Staged for row appending.
  - **`Changed`**: Join key matched; one or more mapped columns differ. Exact Excel coordinates (`CellRef`, e.g. `C14`) and cell diffs are generated.
  - **`Unchanged`**: Join key matched; all mapped fields are identical under configured normalization.
  - **`Ambiguous`**: Join key is null/empty, duplicated in the incoming source, or matches multiple base records. Requires manual operator resolution.
  - **`Removed`**: Base row is missing from the incoming source (when `DetectRemovedRows = $true`).
- **Diff Policy Engine (`DiffPolicy`)**:
  - `TrackChanges` (Default): Standard cell difference tracking for critical attributes (phone, address, car, class).
  - `IgnoreChanges`: Bypasses cell diff generation for existing matched records (suppressing unimportant changes), while keeping projected values for new rows.
  - `NormalizePostalCode`: Neutralizes postal codes (`\b\d{2}-\d{3}\b|\b\d{5}\b`) in city fields before comparison.
  - `FuzzyContainment`: Evaluates bidirectional substring containment (e.g. school name suffix `"w Gliwicach"`), ignoring benign naming extensions.

### Region 6: Metadata Stamping Engine
- Handled by `Get-StampedMetadataValue`.
- Supports dynamic audit tokens stamped into base columns upon writing:
  - `{ChangeDate}`: Current timestamp formatted as `yyyy-MM-dd HH:mm`.
  - `{SourceFileFullPath}`: Absolute path to incoming updates file.
  - `{SourceFileName}`: File name with extension.
  - `{CurrentUser}`: Executing user (`$env:USERNAME`).
  - `{SourceRowNumber}`: Physical row index in incoming file.
  - `{ImportBatchId}`: Unique GUID for the reconciliation session.

### Region 7: Write-Back Engine
- Handled by `Backup-BaseFile`, `Invoke-MasterWriteBack`, and `Write-BaseLogSheet`.
- **Pre-write Backup**: Creates a pristine copy of the base file in `Backups/`.
- **Pruning**: Enforces retention count and 200 MB disk size capping.
- **InPlace Stream Execution**: Calls `[EditExcelHelper]::ApplyCellUpdates()` and `[EditExcelHelper]::AppendRows()`.
- **In-Workbook Audit Sheet**: When `LogToBaseSheet = $true`, appends structured change rows to `ImportLog` worksheet with PII masking support.
- **Atomic Staging**: Writes to `<file>.tmp.xlsx`, validates with `[FastExcelHelper]::ReadSheet`, and replaces the target file via `[System.IO.File]::Replace`.

### Region 8: Audit Logging
- Handled by `Write-ImportLog`.
- Emits three complementary log files:
  1. `import_<yyyy-MM-dd_HHmmss>_<BatchId>.jsonl`: High-precision structured event stream capturing every cell patch and decision.
  2. `import_<yyyy-MM-dd_HHmmss>_<BatchId>.txt`: Formatted human-readable report with tabular statistics.
  3. `masterlog.jsonl`: Cumulative append-only history of all reconciliation sessions.
- Enforces PII anonymization when `RedactNamesInLog = $true`.

### Region 9: WPF UI Engine & Theme System
- Dynamic XAML layout with responsive split-screen cards and resizable rows.
- **Data & Mapping Live Preview Panel (`tabMapping`)**:
  - Resizable bottom card separated by a horizontal `GridSplitter` (`ResizeDirection="Rows"`).
  - Contains `tabsMappingPreview` with three dedicated views:
    1. **⚡ Mapping Result (Preview)** (`gridMappingResultPreview`): Live projection of the current sample row through active mapping rules. Displays Target Base Column, Mapped Source Columns (with merge mode indicator), Projected Value, Current Base Value, and Status badges (`Identical`, `Will Change`, `New Value`, `Unmapped`).
    2. **📥 Incoming File (Sample Row)** (`gridPrevIncomingRow`): Horizontal data row from the incoming file with row number column `#` and sanitized dynamic column keys.
    3. **🗄️ Base File (Sample Row)** (`gridPrevBaseRow`): Horizontal data row from the master base file.
  - **High-Performance Sample Caching**: Fast preview sampling reads only the first 10 rows via `[FastExcelHelper]::ReadSheet($path, $sheet, 10)` in ~11 ms without loading the full workbook into memory.
  - **Sample Row Navigation**: Interactive `◀ Prev Row` and `Next Row ▶` buttons to cycle through sample rows 1 to 10 with live row indicator text (`txtPreviewInfo`).
  - **Automatic Refresh Hooks**: Automatically recalculates preview on file loading, sheet selection change, auto-mapping, manual rule additions/deletions, profile loading, and tab selection.
- **QoL Selective Field Takeover & Inline Card Editing**:
  - **Save as New File (`btnApplyToNewFile`)**: Dedicated action bar button (`💾 Save as New File...`) allowing operators to export all accepted and inline-edited updates directly into a brand new Excel or CSV file without modifying the original master base file.
  - **Direct Inline Card Editing (`txtEditNew`)**: Proposed values can be typed into directly on the review card, setting `$chg.CustomEdited = $true` and updating live staging metrics.
  - **Instant Revert (`↺` / `btnRevert`)**: Restores original proposed file value and clears dirty flags.
  - **Row Tools**: Dedicated `Select All Fields` and `Deselect All Fields` buttons above the diff table for one-click field management on the active record.
  - **Batch Column Toggles (`wrapBatchColToggles`)**: Dynamic column chips to toggle a specific field across all accepted rows session-wide.
  - **Live Staging Summary (`txtStagingSummary`)**: Real-time counter showing `[X staged changes across Y rows (Z skipped)]`.
- **Theme Switching**: On-the-fly toggling between `Dark` (charcoal/slate) and `Light` (clean corporate) palettes without reloading.
- **Keyboard Shortcuts**:
  - `A`: Accept record and advance.
  - `R`: Reject record and advance.
  - `S`: Skip record without decision.
  - `E`: Open modal dialog to edit proposed value.
  - `B` / `Left Arrow` / `Backspace`: Go to previous record.
  - `Ctrl + Z`: Undo last decision.
  - `F1`: Open trilingual interactive Help dialog.

### Region 10: Headless & Batch Mode
- Handled by `Invoke-HeadlessMasterUpdater`.
- Executes automated reconciliations without GUI interaction when invoked with `-Headless` or `-AcceptAllNewAndChanged`.
- Standalone executable cleanly terminates with `exit $ec` to prevent console/process hanging.

---

## 5. Comprehensive Function Index

### Core PowerShell Functions

| Function Name | Region | Parameters | Description |
|---|---|---|---|
| `Get-AppConfig` | 2 | *(none)* | Loads `config.json` or returns default configuration object. |
| `Save-AppConfig` | 2 | `$Config` | Serializes configuration to `config.json` with UTF-8 BOM. |
| `Compute-HeaderFingerprint` | 3 | `$SheetName`, `$Headers` | Calculates SHA-256 fingerprint for layout auto-detection. |
| `Get-MappingProfiles` | 3 | *(none)* | Returns all saved profile JSON objects from disk. |
| `Save-MappingProfile` | 3 | `$Profile` | Persists a column mapping profile to disk. |
| `Build-JoinKey` | 4 | `$Row`, `$KeyColumns`, `$NormalizeOptions` | Constructs composite normalized lookup string. |
| `Get-ProjectedRow` | 4 | `$IncomingRow`, `$MappingRules` | Transforms an incoming row into base schema format. |
| `Invoke-MasterCompare` | 5 | `$BaseRows`, `$IncomingRows`, `$MappingRules`, `$JoinKeys`, `$Options` | Compares files and outputs structured review collection. |
| `Get-StampedMetadataValue` | 6 | `$Token`, `$Context` | Resolves dynamic metadata tokens (`{ChangeDate}`, `{CurrentUser}`, etc.). |
| `Backup-BaseFile` | 7 | `$BaseFilePath`, `$BackupDir`, `$RetentionCount`, `$MaxBackupMb` | Generates pre-write backup and prunes obsolete backups. |
| `Invoke-MasterWriteBack` | 7 | `$BaseFilePath`, `$ReviewItems`, `$Options` | Executes atomic InPlace or SafeRewrite changes to base file. |
| `Write-BaseLogSheet` | 7 | `$BaseFilePath`, `$AuditRecords`, `$SheetName`, `$MaskPii` | Appends change records to dedicated audit sheet in base file. |
| `Write-ImportLog` | 8 | `$LogDir`, `$BatchId`, `$ReviewItems`, `$Stats`, `$MaskPii` | Emits `.jsonl`, `.txt`, and updates `masterlog.jsonl`. |
| `Get-UiString` | 9 | `$Key`, `$Lang` | Returns localized UI string from `language.json`. |
| `Update-Localization` | 9 | `$Lang` | Dynamically updates all XAML text controls and labels. |
| `Export-ReviewReport` | 9 | `$ReviewItems`, `$DestinationPath`, `$BaseFilePath`, `$IncomingPath` | Exports review records to `.xlsx`, `.csv`, or `.html`. |
| `Invoke-HeadlessMasterUpdater` | 10 | `$Params` | Orchestrates headless batch reconciliation without GUI. |
| `Show-MasterUpdater` | 10 | `$Params` | Initializes WPF application window and event loops. |

### Embedded C# Classes & Methods

| Class | Method | Signature | Description |
|---|---|---|---|
| `FastExcelHelper` | `ReadSheet` | `DataTable ReadSheet(string path, string sheetName)` | Fast OpenXML streaming reader; extracts row and cell values into a `DataTable`. |
| `FastExcelHelper` | `GetSheetNames` | `List<string> GetSheetNames(string path)` | Reads `workbook.xml` to list all available worksheet names. |
| `FastExcelHelper` | `ExportToExcel` | `void ExportToExcel(DataTable dt, string path, string sheetName)` | Generates clean OpenXML workbook from `DataTable`. |
| `FastExcelHelper` | `ParseCsv` | `DataTable ParseCsv(string path, char delimiter)` | High-performance RFC-4180 CSV parser handling multiline quotes. |
| `FastDiffHelper` | `Normalize` | `string Normalize(string val, bool ignoreCase, bool trim, bool ignoreSpecial, bool ignoreSpaces)` | Applies string normalization rules for accurate comparison. |
| `FastDiffHelper` | `AreEqual` | `bool AreEqual(string a, string b, DiffOptions opt)` | Compares two cell values under specified normalization flags. |
| `EditExcelHelper` | `ApplyCellUpdates` | `int ApplyCellUpdates(string path, string sheet, List<CellUpdate> updates)` | Streams `sheetN.xml` and patches cell nodes in-place using inline strings. |
| `EditExcelHelper` | `AppendRows` | `int AppendRows(string path, string sheet, List<Dictionary<string, string>> rows)` | Injects new row nodes before `</sheetData>` and updates dimension ref. |
| `EditExcelHelper` | `AppendOrUpdateLogSheet` | `void AppendOrUpdateLogSheet(string path, string sheet, List<LogEntry> entries)` | Creates or appends audit rows to in-workbook log worksheet. |

---

## 6. Localization System & Trilingual Parity

The application enforces a strict **Zero-Hardcoded-Strings** policy. Every UI element, message, dialog title, file filter, and placeholder is retrieved via `Get-UiString`.

### Catalog Architecture (`language.json`)
```json
{
  "DefaultLanguage": "pl",
  "Languages": {
    "en": { "DisplayName": "English", "Strings": { ... 323 keys ... } },
    "pl": { "DisplayName": "Polski",  "Strings": { ... 323 keys ... } },
    "de": { "DisplayName": "Deutsch", "Strings": { ... 323 keys ... } }
  }
}
```

### Invariants & Quality Rules
1. **100% Trilingual Parity**: Any key added to `en` must simultaneously exist in `pl` and `de`. The catalog size must remain identical across all three languages.
2. **Live Dynamic Localization**: Switching language in the GUI (`cmbLanguage`) invokes `Update-Localization`, which iterates over every registered control (`x:Name`) and updates `.Text`, `.Content`, or `.Header` without restarting the process.
3. **Automated Assertion**: `Tests\Test-Localization.ps1` scans `Master-Updater.ps1` for every `Get-UiString 'Key'` invocation and verifies 100% presence across all languages.

---

## 7. Configuration Schema (`config.json`)

| Property | Data Type | Default | Description |
|---|---|---|---|
| `DefaultLanguage` | string | `"pl"` | Active UI language on launch (`en`, `pl`, `de`). |
| `Theme` | string | `"Dark"` | Active color palette (`Dark` or `Light`). |
| `BackupDirectory` | string | `"Backups"` | Directory path for pre-write backup snapshots. |
| `BackupRetentionCount` | integer | `20` | Maximum number of historical backup copies to retain. |
| `MaxBackupMb` | integer | `200` | Disk threshold in megabytes; prunes oldest backups when exceeded. |
| `LogDirectory` | string | `"Logs"` | Directory path for `.jsonl` and `.txt` audit trails. |
| `WriteMode` | string | `"InPlace"` | Write engine mode (`InPlace` or `SafeRewrite`). |
| `RedactNamesInLog` | boolean | `false` | When true, masks names and addresses in log files. |
| `DetectRemovedRows` | boolean | `false` | When true, flags base rows absent from incoming source. |
| `RememberLastBasePath` | boolean | `true` | Persists base file path between sessions. |
| `LastBaseFilePath` | string | `""` | Cached path of the last used master base file. |
| `ShowUnchangedRows` | boolean | `false` | When true, includes unchanged rows in review list by default. |
| `LogToBaseSheet` | boolean | `false` | When true, records audit logs directly in the base workbook. |
| `BaseLogSheetName` | string | `"ImportLog"` | Worksheet name used for in-workbook change logging. |
| `MetadataColumns` | object | *(mappings)* | Configures base audit column names and token bindings. |

---

## 8. Automated Test Suite (12 Test Files)

The test suite provides comprehensive coverage across the entire engine and GUI layer:

1. **`Test-Localization.ps1`**:
   - Asserts 100% trilingual key completeness across `en`, `pl`, and `de` (323 keys each).
   - Validates that all `Get-UiString` invocations in `Master-Updater.ps1` map to valid catalog keys.
2. **`Test-DataMappingPreview.ps1`**:
   - Validates all 22 Data & Mapping Live Preview UI controls in the WPF element tree.
   - Tests sample row extraction from Base and Incoming files, DataGrid binding, and live projection calculation via `Get-ProjectedRow`.
   - Validates status badge classification (`Identical`, `Will Change`, `New Value`, `Unmapped`), row navigation (`◀ Prev` / `Next ▶`), and dynamic localization across EN/PL/DE.
3. **`Test-CompareEngine.ps1`**:
   - Tests composite keys, case-sensitivity, whitespace trimming, and status categorization (`New`, `Changed`, `Unchanged`, `Ambiguous`).
4. **`Test-EditExcelHelper.ps1`**:
   - Validates InPlace XML stream patching, row appending, multi-sheet preservation, and format stability.
5. **`Test-HeadlessAndWatcher.ps1`**:
   - Tests non-interactive command-line execution (`-Headless`, `-AcceptAllNewAndChanged`).
6. **`Test-NegativePath.ps1`**:
   - Tests error handling on locked files, invalid XML, missing worksheets, and corrupted workbooks.
7. **`Test-SampleFilesIntegration.ps1`**:
   - Validates real-world educational fixtures (Students, Grades, Attendance datasets).
8. **`Test-UIComponents.ps1`**:
   - Verifies WPF control instantiation, XAML element finding, and theme brush converter operation.
9. **`Test-UnchangedOption.ps1`**:
   - Tests visibility toggling of unchanged records in review lists, counter calculations, and safety locks on "Accept All".
10. **`Test-BaseSheetLogging.ps1`**:
    - Tests creation and incremental appending of the in-workbook `ImportLog` audit worksheet, header sanitization, and PII masking.
11. **`Test-E2E-MasterUpdater.ps1`**:
    - Full end-to-end integration test: compares files, stages decisions, executes write-back, verifies base file integrity, and inspects audit logs.
12. **`Test-DeepVerification.ps1`**:
    - Rigorous edge-case verification: multi-sheet preservation, formula safety checks, and UTF-8 diacritic fidelity.

---

## 9. Critical Technical Gotchas & Maintenance Rules

### 1. Windows PowerShell 5.1 C# Syntax Restrictions
- **Gotcha**: Windows PowerShell 5.1 compiles C# using the legacy Roslyn/CodeDom compiler (C# 5.0).
- **Rule**: Never use C# 7+ features in `Add-Type` blocks:
  - Do NOT use pattern matching (`if (obj is IEnumerable ie)`). Use `var ie = obj as IEnumerable; if (ie != null)`.
  - Do NOT use string interpolation (`$""`). Use `string.Format()`.
  - Do NOT use out variable declarations (`int.TryParse(s, out int val)`). Declare variable beforehand (`int val; int.TryParse(s, out val)`).

### 2. Parenthetical `(if ...)` Expressions in PowerShell 5.1
- **Gotcha**: Wrapping an `if` statement in parentheses `(if (...) { ... } else { ... })` fails in PowerShell 5.1 with: `The term 'if' is not recognized as the name of a cmdlet`.
- **Rule**: Always assign to a variable first (`$val = if ($cond) { $a } else { $b }`) or wrap in a subexpression `$(if ($cond) { $a } else { $b })`.

### 3. PS2EXE Headless Process Termination
- **Gotcha**: Executables compiled with `PS2EXE` using `-noConsole` will hang in background memory if headless mode simply finishes without explicitly exiting.
- **Rule**: Headless routines must end with `exit $ec` to terminate the process cleanly before the GUI block is evaluated.

### 4. PS2EXE `$MyInvocation.Line` Null Safety
- **Gotcha**: In PS2EXE-hosted environments, `$MyInvocation.Line` is `null`. Evaluating `$MyInvocation.Line.StartsWith('.')` throws `NullReferenceException`.
- **Rule**: Always test `$MyInvocation.Line` for null: `if ($MyInvocation.Line -and $MyInvocation.Line.StartsWith('.'))`.

### 5. Automated Tests & Modal Dialogs
- **Gotcha**: Calling `.RaiseEvent(ClickEvent)` on buttons like `$btnAcceptAll` triggers modal `[System.Windows.Forms.MessageBox]::Show()`, hanging automated test runners indefinitely.
- **Rule**: In headless test scripts, mutate the underlying data model (`$script:AllReviewItems`) directly instead of simulating UI button clicks that open modal message boxes.

### 6. ComboBoxItem `.Tag` Canonical Domain Extraction
- **Gotcha**: Reading `$cmb.SelectedItem` returns a localized `ComboBoxItem` object instead of a string, causing localized text to bleed into SQL parameters or domain models.
- **Rule**: Always store canonical domain keys in `ComboBoxItem.Tag` and display text in `.Content`. Extract `.Tag` when reading selections:
  ```powershell
  $val = if ($cmb.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { $cmb.SelectedItem.Tag } elseif ($cmb.SelectedItem) { $cmb.SelectedItem.ToString() } else { $default }
  ```

### 7. File Encoding Guardrail (UTF-8 with BOM)
- **Gotcha**: PowerShell string replacement tools and standard `Set-Content` in PowerShell 5.1 strip or mangle the UTF-8 BOM (`0xEF, 0xBB, 0xBF`), corrupting non-ASCII characters (Polish diacritics, German umlauts).
- **Rule**: Always write files using explicit UTF-8 with BOM:
  ```powershell
  [System.IO.File]::WriteAllText($path, $content, [System.Text.UTF8Encoding]::new($true))
  ```

### 8. StrictMode Property Resolution on Mapping Rules & Event Scoping
- **Gotcha**: Under `Set-StrictMode -Version Latest`, accessing `$rule.SplitRegex` on a `PSCustomObject` lacking that property throws `PropertyNotFoundException`. Conversely, accessing `.PSObject.Properties['...']` on a `[hashtable]` rule returns `$null`. Furthermore, event handlers registered via `.add_Click({ ... })` execute outside the local scope once `Show-MasterUpdater` returns.
- **Rule**:
  1. Inspect both `IDictionary` (`$rule.Contains('Prop')`) and `PSCustomObject` (`$rule.PSObject.Properties['Prop']`) safely before reading optional rule fields in `Get-ProjectedRow`.
  2. Scope shared UI elements and update routines to `$script:` (e.g. `$script:UpdateDataMappingPreview`) so WPF event delegates can safely invoke them from any dispatch context.

### 9. WPF GridSplitter Layout Rounding, Scrollbar Oscillation & Snap-Back Reset Bug
- **Gotcha**: In WPF, dragging a `GridSplitter` between star rows causes `GridSplitter.MoveSplitter` to verify that `actual1 + actual2` is close to `Original1 + Original2` via `LayoutDoubleUtil.AreClose`. If:
  1. An affected row hits its `MinHeight` (e.g. `MinHeight="180"`), or
  2. A child control (e.g. `DataGrid`) collapses or displays its `ScrollBar`, triggering a layout pass with DPI-scaling sub-pixel discrepancy (e.g. 0.8px rounding on a 6px row at 125% DPI),
  `AreClose` returns `false`, causing WPF to invoke `GridSplitter.CancelResize()` and snap all rows back to default star heights mid-drag. Furthermore, excessive `MinHeight` prevents collapsing panels.
- **Rule**:
  1. Align splitter row height to clean integer device-pixel multiples (use `Height="8"`, not `6`).
  2. Reduce `MinHeight` constraints to the absolute minimum necessary (e.g. `MinHeight="48"` for the preview header, `MinHeight="90"` for mapping rules) so users can freely collapse or expand panels.
  3. Specify explicit `ResizeBehavior="PreviousAndNext"` and `ResizeDirection="Rows"`.
  4. Provide a visible grab handle template (`ControlTemplate` with subtle divider line and hover pill) so users have clear visual affordance.

### 10. Array Coercion & Pipeline Unrolling under PowerShell 5.1 StrictMode
- **Gotcha**: In Windows PowerShell 5.1, when an `if` expression returns a collection, PowerShell's pipeline automatically unrolls single-element collections into a scalar `PSCustomObject`, and empty collections into `$null`. Under `Set-StrictMode -Version Latest`, `[PSCustomObject]` does not possess a `.Count` property (unlike PS 7+ ETS). Evaluating `$matchList.Count` throws:
  `PropertyNotFoundException: The property 'Count' cannot be found on this object. Verify that the property exists.`
- **Rule**:
  Always force array coercion on collection-returning expressions:
  ```powershell
  $matchList = @(if ($baseIndex.ContainsKey($uKey)) { $baseIndex[$uKey] } else { })
  $comp      = @(Invoke-MasterCompare ...)
  ```

### 11. Resilient Legacy CSV Encodings (Windows-1250 / ANSI) & Pipe Delimiters
- **Gotcha**: Legacy Polish ERP/accounting exports (Symfonia, Comarch Optima, Subiekt GT, SAP dumps) frequently produce pipe-delimited (`|`) or semicolon-delimited CSVs encoded in Windows-1250 without a Byte Order Mark. Standard `StreamReader(path, Encoding.UTF8, true)` falls back to UTF-8, mangling Polish diacritics into replacement characters (``).
- **Rule**:
  1. Inspect stream headers for BOMs (UTF-8, UTF-16 LE, UTF-16 BE).
  2. If absent, scan initial bytes for valid UTF-8 multibyte sequences. If invalid bytes are detected in the `0x80-0xFF` range, automatically fall back to Windows-1250 (`Encoding.GetEncoding(1250)`).
  3. Include pipe `|` alongside `;`, `,`, and `\t` in automatic delimiter frequency analysis (`DetectCsvDelimiter`).

### 12. Trilingual Documentation Parity & Comment-Based Help (CBH) Completeness
- **Gotcha**: Adding new CLI parameters (`-SummaryJsonPath`, `-AutoAccept`, `-Headless`) or UI features without synchronizing in-app help (`F1`), localized catalogs, and standalone guides leads to documentation drift and user confusion.
- **Rule**:
  1. Every new or updated CLI parameter must be documented in script-level CBH and function CBH (`Invoke-HeadlessMasterUpdater`).
  2. Maintain 100% key and text parity across English (`USER_GUIDE_EN.md`), Polish (`USER_GUIDE_PL.md`), German (`USER_GUIDE_DE.md`), and consolidated `USER_GUIDE.md`.
  3. Wire in-app help modal cards directly into `language.json` (`HelpWorkflowCard*`, `HelpShortcutsCard*`).
  4. Enforce UTF-8 with BOM across all documentation and configuration artifacts via `[System.IO.File]::WriteAllText($path, $content, [System.Text.UTF8Encoding]::new($true))`.

### 13. Multi-Column Concatenate Mapping & N:1 / 1:N Column Transformations
- **Gotcha**: When matching records between Excel sheets where one file stores address data in a single combined cell (e.g. `Adres: "Ulica 10, Miasto"`) and the other stores address data in separate columns (e.g. `Adres: "Ulica 10"`, `Miasto: "Miasto"`), comparing individual columns causes false differences. Furthermore, standard single-selection ComboBoxes in mapping dialogs prevent selecting multiple source or target columns.
- **Rule**:
  1. Provide multi-select ListBoxes (`SelectionMode="Extended"`) for both Base and Incoming columns in the Rule Editor dialog (`ShowRuleDialog`).
  2. In `Invoke-MasterCompare`, for any rule configured with `MergeMode = 'Concatenate'` and multiple columns on either side, evaluate overall merged equality first via `[FastDiffHelper]::MergeValues` and `[FastDiffHelper]::AreEqual`. If merged strings are equal, bypass individual column diffs to avoid false positives.
  3. In `Get-ProjectedRow`, automatically split incoming values across multiple base columns when `MergeMode = 'Concatenate'` and `Separator` is specified (e.g. `", "`), ensuring cell write-back accurately updates each individual base cell.
  4. Provide both "+ Add Rule" and "Edit Rule" capabilities (including double-click on `gridMappingRules`) with live preview cards displaying formatted Concatenate combinations.

### 14. Comparison Mode Flags & Scope-Safe Dynamic Localization Dispatch
- **Gotcha**:
  1. Comparing text data without flexible matching flags causes false discrepancies due to case differences, leading/trailing whitespace, punctuation differences (e.g. `.` vs `,` in addresses), or multi-spaces.
  2. In WPF PowerShell applications under `Set-StrictMode`, event callbacks (such as language selection changes or button clicks) execute outside the local scope of `Show-MasterUpdater`. Accessing local control variables directly or invoking `$ApplyTheme` / `$UpdateLocalization` throws `PropertyNotFoundException: The variable '$chkTrimWhitespace' cannot be retrieved because it has not been set` or `The property 'Background' cannot be found on this object`.
- **Rule**:
  1. Expose 4 comparison checkboxes matching `Compare-ExcelFiles`:
     - *Ignore case when comparing* (`chkIgnoreCase`)
     - *Trim whitespace before comparing* (`chkTrimWhitespace`)
     - *Ignore punctuation & special characters* (`chkIgnoreSpecialChars`)
     - *Ignore internal whitespace differences* (`chkIgnoreAllSpaces`)
  2. Implement high-performance matching in `[FastDiffHelper]::AreEqual` using compiled regular expressions (`_rgxSpecial = [^\p{L}\p{Nd}\s]` and `_rgxWhitespace = \s+`).
  3. Scope shared UI elements to `$script:` (e.g. `$script:chkIgnoreCase`, `$script:chkTrimWhitespace`, etc.) and resolve dynamic UI elements in `$UpdateLocalization` and `$ApplyTheme` safely through `$w.FindName($name)` with null-guards and error handling.
  4. Persist and restore comparison mode options in mapping profiles (`ProfileStore`) and `AppConfig`, while immediately triggering live preview recalculation (`& $script:UpdateDataMappingPreview`) upon any checkbox toggle.

### 15. Trilingual In-App Help System & Smart Auto-Mapping Heuristics
- **Gotcha**:
  1. Users and operators often struggle with complex column relationships (e.g. `Exact` vs `FirstNonEmpty` vs `Concatenate`) without context-rich in-app documentation.
  2. ERP file exports frequently contain subtle header typos (`pywyżej/ ponieżej`) or minor abbreviations (`niepełn.`), causing standard exact-string AutoMap to fail and leave columns unmapped.
  3. Real-world files often split address and city into two columns in the base database while providing a single combined address in monthly change files.
- **Rule**:
  1. Maintain a dedicated 4th tab (`HelpTabMergeModes`) in `ShowHelpDialog` (F1) explaining Merge Modes, 4 Comparison Options, and Smart Auto-Mapping heuristics across English, Polish, and German.
  2. In `Invoke-AutoMapRules`, implement `Find-SmartHeaderMatch` using punctuation stripping and fuzzy word-stem matching (>= 50% overlap) to automatically link misspelled or abbreviated ERP headers.
  3. In `Invoke-AutoMapRules`, detect when Base has separate Address and City columns while Incoming has a single combined Address; automatically generate a `Concatenate` rule with `, ` separator.
  4. Ensure 100% key and text parity in `language.json` (323 keys) and all standalone guides (`USER_GUIDE_PL.md`, `USER_GUIDE_EN.md`, `USER_GUIDE_DE.md`, `USER_GUIDE.md`) with UTF-8 BOM encoding.

### 16. UI Modernization, Metadata Settings, and Resizing Grids
- **Gotcha**:
  1. Traditional Windows scrollbars look archaic in modern Dark/Light theme applications, clashing with the custom WPF palette.
  2. Fixed column widths in the Review Details pane limit readability when inspecting rows with numerous mapped columns or long text blocks.
  3. Stamping dynamic tracking metadata (e.g., `ChangeDate`, `SourceFileName`) requires specific configuration to map the runtime tokens into physical base columns.
- **Rule**:
  1. Inject modern, flat XAML `ScrollBar` and `Thumb` templates directly into `Window.Resources`, bound to the active theme palette (`BgCardHover`, `BorderCard`).
  2. Implement a vertical `GridSplitter` inside the Review Item details pane (between properties list and cell diff grid) allowing dynamic horizontal resizing. Enable `TextWrapping="Wrap"` inside `DataGrid` cells.
  3. Expose a dedicated **Metadata Columns (Kolumny metadanych)** tab in the Settings modal to allow users to bind `ModelName`, `ChangeDate`, `SourceFileFullPath`, `SourceFileName`, `CurrentUser`, `SourceRowNumber`, and `ImportBatchId` without cluttering the primary mapping dialog.
  4. Document Metadata Tokens in a dedicated in-app Help Dialog tab (`HelpTabSettingsMetadata`).

### 17. InPlace OpenXML Style, Table & Color Preservation vs Destructive SafeRewrite
- **Gotcha**:
  1. **Destructive SafeRewrite Fallback**: Previously, when an InPlace guard warning was raised (e.g. for `tableParts` or shared formulas), `Invoke-MasterWriteBack` automatically fell back to `SafeRewrite`. `SafeRewrite` called `[FastExcelHelper]::ExportToExcel`, which generated a plain OpenXML workbook from scratch, permanently destroying 100% of cell colors, font formatting, borders, conditional formatting rules, Excel tables, and formulas.
  2. **PowerShell AppDomain Assembly Locking**: In PowerShell, types compiled via `Add-Type` (such as `EditExcelHelper`) cannot be modified or re-compiled within the same running process. Dot-sourcing `. .\start.ps1` in an active session retained the old cached C# type with legacy guard warnings.
  3. **Table Range Desynchronization on Append**: Appending new rows directly into `<sheetData>` without expanding `xl/tables/table*.xml` left appended records outside the table boundary, losing table styles (such as zebra stripes) and auto-filter integration. Furthermore, newly appended cells lacked `s` (style) attributes.
- **Rule**:
  1. **Strict InPlace Guarantee for Excel Workbooks**: `.xlsx` files must **ALWAYS** use `InPlace` OpenXML patching (`EditExcelHelper.WriteChanges`), never falling back to `SafeRewrite` or `ExportToExcel`. `SafeRewrite` is strictly restricted to plain `.csv` files.
  2. **Automatic Excel Table Expansion**: `EditExcelHelper.WriteChanges` scans for any `xl/tables/table*.xml` packages, expanding `ref` and `autoFilter.ref` (e.g., `A1:AL493` -> `A1:AL494`) when rows are appended so new records seamlessly inherit table zebra striping and column filters.
  3. **Style Inheritance for Appended Cells**: When appending new rows, cells copy the style attribute `s` from the corresponding column of the preceding data row.
  4. **Clean Session Relaunch in `start.ps1`**: `start.ps1` checks if `EditExcelHelper` is already loaded in the process and automatically relaunches in a clean `-NoProfile -STA` process to prevent stale in-memory assembly conflicts.

### 18. PS2EXE Standalone Executable Startup Path Binding & StrictMode Resilience
- **Gotcha**:
  1. **Empty `$PSScriptRoot` in Hosted Runspaces**: Inside executables generated by `Invoke-ps2exe`, scripts run within an in-memory runspace where `$PSScriptRoot` is an empty string `""` (not `$null`), and `$MyInvocation.MyCommand.Path` is `$null` or `""`.
  2. **ParameterBindingValidationException on Startup**: Calling `Split-Path -Parent $PSScriptRoot` or `Join-Path $PSScriptRoot ...` when `$PSScriptRoot` is `""` throws: `Cannot bind argument to parameter 'Path' because it is an empty string.` This triggered an unhandled exception dialog during `Import-LanguageCatalog` before WPF initialization.
  3. **VariableIsUndefined in StrictMode**: Referencing `$ScriptRoot` or undeclared `$script:AppRootDir` directly under `Set-StrictMode -Version Latest` throws `VariableIsUndefined`.
  4. **PS2EXE `-embedFiles` Parameter Format**: In `ps2exe` v1.0.18+, `-embedFiles` requires a `[hashtable]` (`@{'TargetPath' = 'SourcePath'}`). Passing an array (`@($file)`) causes `ps2exe` to skip embedding files entirely.
- **Rule**:
  1. **Universal Safe Root Directory Helper**: Centralize directory resolution in `Get-ApplicationRootDir` using `Test-Path 'variable:...'` to safely probe `$PSScriptRoot`, `$ScriptRoot`, `[System.AppDomain]::CurrentDomain.BaseDirectory`, and `$MyInvocation.MyCommand.Path` without parameter binding or strict mode errors.
  2. **Guarded Candidate Resolution**: Build directory candidate lists dynamically without evaluating `Join-Path` or `Split-Path` on empty or null strings.
  3. **Proper `-embedFiles` Packaging**: In `Build-Exe.ps1`, pass a hashtable for `-embedFiles` (`@{' .\language.json' = $langFile; '.\Docs\...' = ...}`) to guarantee self-contained extraction alongside the compiled executable.
