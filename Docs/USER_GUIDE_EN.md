# Excel Master Updater — English User Guide

## 1. Overview & Purpose

**Excel Master Updater** is a desktop application designed to maintain and synchronize a single **Master ("Base") Excel file** of records (e.g. students, inventory, employees) from recurring incoming update files.

### Key Capabilities
- **Universal Column Mapping**: Map arbitrary incoming column names to base fields once; subsequent files with the same header structure are auto-detected via template fingerprint.
- **Row-by-Row Review**: Inspect every proposed change in a side-by-side visual diff card with per-cell toggle checkboxes and double-click quick editing.
- **InPlace OpenXML Patching**: Updates only modified cells directly in the worksheet XML stream without touching cell colors, fonts, column widths, or other worksheets.
- **Resilient CSV Engine**: Automatically detects delimiters (semicolon `;`, comma `,`, tab `\t`, pipe `|`) and character encodings (UTF-8 with/without BOM, UTF-16, and Windows-1250 Central European ANSI fallback).
- **Automatic Audit Stamping**: Stamps change date, source file name, username, source row number, and session GUID directly into base audit columns.
- **Automated Pre-Write Backups & Summaries**: Generates timestamped backups in `Backups/` before any write operation, and exports structured execution summary JSON metrics for automated pipelines.

---

## 2. Quick Start

### Launching the Application
Launch via PowerShell 7 (recommended):
```powershell
pwsh -File .\start.ps1
```
Or via Windows PowerShell 5.1:
```powershell
powershell -File .\start.ps1
```
Or run the standalone windowed binary directly:
```powershell
.\Master-Updater.exe
```

### Pre-Loading Files via Command Line
```powershell
pwsh -File .\start.ps1 -BaseFilePath "C:\Data\Master.xlsx" -IncomingPath "C:\Data\Updates_September.xlsx"
```

### Unattended Headless Batch Execution
Run automated synchronization without launching the graphical interface:
```powershell
pwsh -File .\Master-Updater.ps1 -BaseFilePath "C:\Data\Master.xlsx" -IncomingPath "C:\Data\Updates.csv" -Headless -AutoAccept AllNonAmbiguous -ExportReportPath "C:\Reports\diff.html" -SummaryJsonPath "C:\Logs\summary.json"
```

---

## 3. Workflow Steps

### Step 1: File Selection & Fingerprint Auto-Mapping
1. **Select Base File**: Choose your master Excel file (`.xlsx`) or delimited `.csv`. Select the target worksheet (e.g. `Students`).
2. **Select Incoming Updates File**: Drag-and-drop or browse for the incoming updates file (`.xlsx` or `.csv`). Delimiters (`;`, `,`, `\t`, `|`) and encodings are detected automatically. Select the source worksheet.
3. **Template Auto-Detection**:
   - If this template has been mapped before, the application recognizes its SHA-256 header fingerprint and automatically loads the saved profile.
   - If this is a new template, click **⚡ Auto Map** or **+ Add Rule** to define mappings.
4. **Mapping Rules Search & Filter**:
   - In the mapping table toolbar, type in the **Search mapping rules...** box (`txtSearchMapping`) to filter rules in real time by target base column or incoming source columns. Click the **`✖`** button to clear the filter.
5. **Join Key Identification & Live Indicator**:
   - In the Join Key lists, select the column(s) that identify uniqueness (e.g. `Id` in base and `StudentId` in incoming, or composite keys like `FirstName` + `LastName`).
   - Changing join key selection immediately refreshes the **Data & Mapping Preview** table below, and marks selected join keys with a prominent key icon (**`🔑 `**) in the Target Base Column.
6. **Live Data & Mapping Preview (Bottom Card)**:
   - Inspect sample rows and test your mapping rules in real time at the bottom of the page before running comparison:
     - **⚡ Mapping Result (Preview)**: Displays all base columns, their mapped source columns, the live projected value from the incoming row, the current base value, and an instant status badge (`Identical`, `Will Change`, `New Value`, or `Unmapped`).
     - **📥 Incoming File (Sample Row)**: Displays the raw sample row from the incoming update file.
     - **🗄️ Base File (Sample Row)**: Displays the raw sample row from the master base file.
   - Use **`◀ Prev Row`** and **`Next Row ▶`** to inspect different sample records (rows 1 to 10).
   - Click **`🔄 Refresh Preview`** to force a fresh read from disk.
7. Click **Start Comparison ▶** to execute the matching engine.

---

### Step 2: Review & Approval

Upon completion of the comparison, the application automatically switches to the **2. Review & Approval** tab.

```
+-----------------------------------------------------------------------------------------------+
| Filter: [ All (50) ] [ New (2) ] [ Changed (3) ] [ Ambiguous (0) ]                            |
+------------------------------------+----------------------------------------------------------+
| #1 D-1001 Smith John     [NEW]     |  StudentId: D-1001                                       |
| #2 D-1002 Doe Jane       [CHANGED] |  FullName: Doe Jane                                      |
| #3 D-1003 Brown Peter    [CHANGED] |  City: London                                            |
|                                    |  Address: High Street 10                                 |
+------------------------------------+----------------------------------------------------------+
| [◀ Back (B)] [✔ Accept (A)] [✖ Reject (R)] [⏭ Skip (S)] [✏ Edit value (E)]                     |
| [✔✔ Accept All] [✖✖ Reject All] [📊 Export Report...]             [ Apply accepted changes ]  |
+-----------------------------------------------------------------------------------------------+
```

#### Status Categories
| Status | Badge Color | Description |
|---|---|---|
| **NEW** | Green (`#064E3B`) | Key does not exist in master file. Row will be appended to base. |
| **CHANGED** | Amber (`#713F12`) | Key matched. Differs in one or more mapped columns. |
| **AMBIGUOUS** | Blue (`#1E3A8A`) | Multiple matches, duplicate keys in source, or empty key. Requires manual candidate selection. |
| **REMOVED** | Red (`#7F1D1D`) | Base record missing from incoming file (when `DetectRemovedRows` is enabled). |
| **UNCHANGED** | Grey (`#1F2937`) | Key matched; all mapped column values identical. Auto-skipped by default; toggle "Show unchanged" or filter by "Unchanged" to review and inspect base values. |

#### Keyboard & Mouse Shortcuts
| Shortcut | Action | Description |
|---|---|---|
| **`A`** | **Accept** | Marks record as accepted and advances to next row. |
| **`R`** | **Reject** | Marks record as rejected and advances to next row. |
| **`S`** | **Skip** | Skips record without decision. Will reappear on next import. |
| **`E`** | **Edit** | Opens dialog to manually override a proposed value before writing. |
| **Double-Click** | **Quick Edit** | Double-clicking any record in the review items list directly opens the override dialog. |
| **`B`** / **`←`** | **Back** | Navigates back to the preceding row. |
| **`Ctrl + Z`** | **Undo** | Reverts the last decision made. |
| **`F1`** | **Help** | Opens the interactive in-app help modal. |

#### Selective Cell Updates
For changed records, the detail pane presents every modified field in a side-by-side card with an **Apply checkbox**. You can uncheck individual cells to selectively retain the current base value while accepting other field updates.

#### Batch Operations & Filtering
- **Show Unchanged Toggle & Filter**: Check "Show unchanged" next to the filter dropdown or choose "Unchanged" in the status filter to display identical rows in the review list without re-running comparison.
- **✔✔ Accept All**: Marks all non-ambiguous rows as accepted in one click (skips unchanged records to protect existing values).
- **✖✖ Reject All**: Marks all rows as rejected.
- **📊 Export Report**: Exports the entire review summary table to formatted Microsoft Excel (`.xlsx`), standalone responsive HTML report (`.html`), or delimited `.csv`.

---

### Step 3: Write-Back Execution

1. When at least one record has been accepted, the **Apply accepted changes to base** button activates.
2. Click the button to open the confirmation dialog.
3. The engine creates a pre-write backup in `Backups/`, appends missing metadata headers to row 1, applies cell modifications in-place, and appends new rows.
4. Generates audit logs in `Logs/` (`.jsonl` and `.txt`), exports execution summary metrics (`last_execution_summary.json` or custom `-SummaryJsonPath`), and optionally appends change audit rows directly to a dedicated sheet (default: `ImportLog`) inside the base workbook.
5. Displays a success notification with update counts and audit sheet confirmation.

---

## 4. Column Merge Modes & Comparison Options

The application includes an advanced data transformation and normalization engine designed to reconcile structural differences across ERP exports, school registers, and HR spreadsheets.

### Column Mapping Merge Modes

Each mapping rule in the table operates in one of three configurable merge modes:

1. **Exact (1:1)**:
   - Direct 1-to-1 mapping from a single incoming column to a single base column.
   - Triggers a change recommendation whenever the incoming value differs from the current base value.

2. **FirstNonEmpty (N:1)**:
   - Provides fallback coalescing across multiple candidate incoming columns.
   - Enables selecting several source columns (e.g. `MobilePhone`, `HomePhone`, `WorkPhone`).
   - The target base column is populated with the value from the first column that contains non-blank text.

3. **Concatenate (N:1 and 1:N)**:
   - **N:1 Joining (Multiple incoming columns into one base column)**:
     - Joins multiple source fields (e.g. `StreetAddress` + `PostalCode` + `City`) into a unified base column `FullAddress` using a configurable delimiter (e.g. `, ` or space).
   - **1:N Splitting (One incoming column across multiple base columns)**:
     - When incoming files combine address and city into a single string (e.g. `10 High Street, London`), but the base file stores them in separate columns (`Address` and `City`), the engine splits the incoming text according to the specified delimiter.
   - **Composite Equality (Zero False Positives)**:
     - Before flagging differences, the comparison engine verifies composite equality. If joining the base columns reproduces the incoming combined text, the row is marked as `Unchanged` with zero false change warnings.

---

### Comparison & Text Normalization Options

The mapping table header provides 4 independent normalization switches:

| Option | Behavior | Example |
|---|---|---|
| **Ignore Case** | Evaluates uppercase and lowercase strings as equal. | `"LONDON"` == `"London"` |
| **Trim Whitespace** | Strips leading and trailing spaces prior to comparison. | `" John "` == `"John"` |
| **Ignore Punctuation & Special Chars** | Compares alphanumeric content while ignoring dots, commas, hyphens, and slashes. | `"St. John St-10"` == `"St John St 10"` |
| **Ignore Internal Whitespace** | Collapses consecutive spaces, tabs, and line breaks into a single space before comparison. | `"Smith   John"` == `"Smith John"` |

---

### Smart Auto-Mapping Heuristics

Clicking **⚡ Auto Map** invokes a multi-tiered heuristic engine:
1. **Exact & Case-Insensitive Matching**: Matches identical column headers.
2. **Smart Fuzzy Stem Matching**:
   - Strips punctuation and matches header word stems (requiring >= 50% stem overlap).
   - Tolerates common typos (e.g. `pywyżej` ➔ `powyżej`) and ERP abbreviations (e.g. `niepełn.` ➔ `niepełno.`, `Adres zakładu pracy` ➔ `Adres Pracy`).
3. **Automatic Split Address & City Detection**:
   - When the base file contains separate columns for address and city while incoming data has a single combined address column, Auto-Map automatically configures a **Concatenate** rule with `, ` separator.
4. **SHA-256 Fingerprint Profiles**:
   - The column mapping configuration is remembered and automatically reapplied on future imports of matching header layouts.

---

## 5. Settings & Metadata Configuration

Click **⚙ Settings** in the top toolbar to configure:

### General Tab
- **Backup Retention Count**: Number of historical backup files to preserve (default: 20). Older backups are automatically pruned.
- **Write Mode**:
  - `InPlace` (default): Fast OpenXML stream patching. Preserves styles and multi-sheet structure.
  - `SafeRewrite`: Full workbook rewrite via FastExcelHelper.
- **Mask PII in logs**: Replaces personal names and addresses with masked tokens (`***`) in log files and the base audit sheet.
- **Mark missing records as removed**: Automatically detects base records absent from the incoming file and stages `Status = 'Removed'`.
- **Show unchanged records by default**: Automatically enables the display of unchanged rows in the review list after comparison.
- **Log changes to a sheet in the base file**: When enabled, records every cell change, added row, and removed record directly into an audit worksheet inside the base Excel workbook.
- **Log sheet name in base file**: Custom name for the in-file audit sheet (default: `ImportLog`, max 31 characters; invalid characters `\ / ? * : [ ]` are automatically sanitized).

### Metadata Columns Tab
Configures dynamic audit columns stamped on every accepted write:
| Base Column | Default Token | Description | Format |
|---|---|---|---|
| `LastModified` | `ChangeDate` | Date and time of update | `yyyy-MM-dd HH:mm` |
| `SourcePath` | `SourceFileFullPath` | Full path of the incoming file | *(none)* |
| `SourceFile` | `SourceFileName` | File name of the incoming file | *(none)* |
| `ModifiedBy` | `CurrentUser` | Username executing the update | `$env:USERNAME` |
| `SourceRow` | `SourceRowNumber` | Physical row number in incoming file | *(none)* |
| `ImportId` | `ImportBatchId` | Unique session GUID | *(none)* |

---

## 6. Emergency Recovery & Restore

### Restore from Backup Toolbar Action
If an erroneous write occurs:
1. Ensure the base file path is selected in the top bar.
2. Click **↺ Restore Backup** in the top toolbar.
3. The application displays the most recent timestamped backup file and asks for confirmation.
4. Upon confirmation, the base file is atomically restored from the backup and reloaded into the UI.

### Crash Recovery
If a power interruption or system crash occurs during writing:
- The engine uses temporary atomic staging files (`<file>.tmp.xlsx`).
- On subsequent launch, the application automatically detects orphaned `.tmp.xlsx` files and prompts to clean up temporary files and restore from the latest backup.

---

## 7. Backups and Audit Logs Location

### Where are the Backup Files Saved?
- **Folder Location**: Stored in the `Backups/` directory located directly inside the application installation folder (e.g. `UpdateExcelBaseFileProject\Backups\`).
- **Config Option**: Can be customized in `AppConfig.BackupDirectory` or inspected via **⚙ Settings**.
- **Quick Access**: Click the **📁 Backups** button in the top toolbar to open the backup directory directly in Windows File Explorer.
- **File Naming Format**: `<BaseFileName>.<yyyy-MM-dd_HHmmss>.bak.xlsx` (e.g. `Master_Students.2026-10-03_233000.bak.xlsx`).
- **Retention & Size Limits**:
  - Automatically retains the last **20** backups by default (configurable under **⚙ Settings**).
  - Automatically prunes older backups when the folder size exceeds **200 MB** (`MaxBackupMb`).

### Where is the Log File Saved?
- **Folder Location**: Stored in the `Logs/` directory located directly inside the application installation folder (e.g. `UpdateExcelBaseFileProject\Logs\`).
- **Config Option**: Can be customized in `AppConfig.LogDirectory` or inspected via **⚙ Settings**.
- **Quick Access**: Click the **📋 Logs** button in the top toolbar to open the log directory directly in Windows File Explorer.
- **Log Formats & Files Generated**:
  1. `import_<yyyy-MM-dd_HHmmss>_<BatchId>.jsonl`: Detailed structured audit trail of every row decision, cell patch, and metadata modification in JSON-Lines format.
  2. `import_<yyyy-MM-dd_HHmmss>_<BatchId>.txt`: Formatted human-readable text audit report with summary statistics and tabular list of accepted/rejected/skipped rows.
  3. `masterlog.jsonl`: Cumulative running ledger appending every session's execution summary.
- **PII Protection**: If **Mask PII in logs** is enabled in Settings, student names, phone numbers, and addresses are automatically redacted from log files.

---

## 8. Tooltips, Shortcuts & In-App Help

- **Interactive In-App Help**: Press **`F1`** at any time or click **❓ Help** in the top toolbar to open the interactive help dialog with full workflow guides, merge modes, comparison options, folder shortcuts, and keyboard references.
- **Comprehensive Tooltips**: Hover over any button, input card, dropdown, or table control to view a descriptive tooltip explaining its action and associated keyboard shortcuts.
- **Multi-Language Support**: All tooltips, help screens, and user guides are fully localized in **English**, **Polish (Polski)**, and **German (Deutsch)**.
