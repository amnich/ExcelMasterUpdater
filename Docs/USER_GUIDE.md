# Excel Master Updater — Documentation / Dokumentacja / Dokumentation

Choose language / Wybierz język / Sprache auswählen:
- 🇬🇧 **[English User Guide](#excel-master-updater--english-user-guide)** (or standalone file: [USER_GUIDE_EN.md](USER_GUIDE_EN.md))
- 🇵🇱 **[Polska Instrukcja Obsługi](#excel-master-updater--polska-instrukcja-obsługi)** (lub osobny plik: [USER_GUIDE_PL.md](USER_GUIDE_PL.md))
- 🇩🇪 **[Deutsches Benutzerhandbuch](#excel-master-updater--deutsches-benutzerhandbuch)** (oder eigenständige Datei: [USER_GUIDE_DE.md](USER_GUIDE_DE.md))

---

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

---

# Excel Master Updater — Polska Instrukcja Obsługi

## 1. Przeznaczenie i kluczowe możliwości

**Excel Master Updater** to aplikacja okienkowa przeznaczona do utrzymywania i precyzyjnej aktualizacji **głównego pliku bazy Excel** na podstawie cyklicznie spływających plików zmian.

### Kluczowe możliwości
- **Uniwersalne mapowanie kolumn**: Skonfiguruj mapowanie nagłówków raz; kolejne pliki o tym samym układzie kolumn są rozpoznawane automatycznie za pomocą unikalnego odcisku SHA-256.
- **Przegląd zmian wiersz po wierszu**: Weryfikacja każdej proponowanej zmiany na czytelnych kartach wizualnych z możliwością odznaczenia pojedynczych komórek oraz szybką edycją dwuklikiem.
- **Bezpieczny silnik InPlace OpenXML**: Aktualizuje wyłącznie zmienione komórki bezpośrednio w strumieniu XML arkusza, bez naruszania kolorów komórek, krojów czcionek, szerokości kolumn i pozostałych arkuszy.
- **Odporny silnik CSV**: Automatyczne wykrywanie separatorów (średnik `;`, przecinek `,`, tabulator `\t`, potok `|`) oraz stron kodowych (UTF-8 z BOM i bez BOM, UTF-16, a także Windows-1250 ANSI).
- **Automatyczne stemple audytowe**: Wstawia do bazy datę zmiany, nazwę pliku źródłowego, login użytkownika, fizyczny numer wiersza i unikalny identyfikator wsadu GUID.
- **Automatyczne kopie zapasowe i raporty JSON**: Tworzy timestampowane kopie w katalogu `Backups/` przed każdą modyfikacją oraz generuje maszynowe podsumowania JSON do integracji CI/CD.

---

## 2. Szybki start

### Uruchamianie aplikacji
Uruchomienie w środowisku PowerShell 7 (zalecane):
```powershell
pwsh -File .\start.ps1
```
Lub w Windows PowerShell 5.1:
```powershell
powershell -File .\start.ps1
```
Lub bezpośrednie uruchomienie pliku wykonywalnego:
```powershell
.\Master-Updater.exe
```

### Wstępne ładowanie plików z wiersza poleceń
```powershell
pwsh -File .\start.ps1 -BaseFilePath "C:\Dane\Baza_Glowna.xlsx" -IncomingPath "C:\Dane\Zmiany_Wrzesien.xlsx"
```

### Automatyczne przetwarzanie w tle (Headless)
Uruchomienie w pełni automatycznej synchronizacji bez otwierania interfejsu graficznego:
```powershell
pwsh -File .\Master-Updater.ps1 -BaseFilePath "C:\Dane\Baza_Glowna.xlsx" -IncomingPath "C:\Dane\Zmiany.csv" -Headless -AutoAccept AllNonAmbiguous -ExportReportPath "C:\Raporty\diff.html" -SummaryJsonPath "C:\Logi\summary.json"
```

---

## 3. Kroki procedury synchronizacji

### Krok 1: Wybór plików i mapowanie kolumn
1. **Wybór pliku bazy**: Wybierz główny plik Excel (`.xlsx`) lub plik tekstowy `.csv`. Wybierz docelowy arkusz danych.
2. **Wybór pliku zmian**: Przeciągnij i upuść lub wskaż plik z nowymi danymi (`.xlsx` lub `.csv`). Separatory (`;`, `,`, `\t`, `|`) oraz strona kodowa wykrywane są automatycznie. Wybierz arkusz źródłowy.
3. **Automatyczne rozpoznawanie szablonu**:
   - Jeśli nagłówki były już wcześniej mapowane, aplikacja automatycznie rozpozna szablon (odcisk SHA-256) i wczyta zapisany profil.
   - Jeśli to nowy szablon, kliknij **⚡ Automatyczne mapowanie** lub **+ Dodaj regułę**, aby przypisać kolumny.
4. **Wyszukiwanie i filtrowanie reguł mapowania**:
   - W pasku narzędzi tabeli mapowania wpisz szukaną frazę w polu **Filtruj reguły mapowania...** (`txtSearchMapping`), aby w czasie rzeczywistym odfiltrować kolumny bazy lub pliku zmian. Kliknij **`✖`**, aby wyczyścić filtr.
5. **Identyfikacja klucza złączenia i wskaźnik na żywo**:
   - Na listach kolumn klucza złączenia wskaż pole unikalne (np. `Identyfikator` w bazie i `KodUcznia` w pliku zmian, lub klucz złożony, np. `Nazwisko` + `Imie`).
   - Zmiana zaznaczenia kluczy natychmiast odświeża tabelę **Podglądu danych i mapowania**, oznaczając wybrane klucze symbolem klucza (**`🔑 `**) w kolumnie docelowej bazy.
6. **Podgląd danych i wyników mapowania na żywo (Dolna sekcja)**:
   - Weryfikuj przykładowe wiersze oraz działanie reguł mapowania w czasie rzeczywistym u dołu ekranu przed rozpoczęciem porównania:
     - **⚡ Wynik mapowania (Podgląd)**: Pokazuje wszystkie kolumny bazy, przypisane kolumny źródłowe, obliczoną wartość wynikową, bieżącą wartość w bazie oraz etykietę statusu (`Identyczne`, `Zmiana`, `Nowa wartość` lub `Niezmapowana`).
     - **📥 Plik zmian (Wiersz próbki)**: Wyświetla poziomy wiersz próbki z pliku zmian.
     - **🗄️ Plik bazy (Wiersz próbki)**: Wyświetla poziomy wiersz próbki z pliku głównego.
   - Użyj przycisków **`◀ Poprzedni wiersz`** i **`Następny wiersz ▶`**, aby przeglądać kolejne rekordy próbki (wiersze 1–10).
   - Kliknij **`🔄 Odśwież podgląd`**, aby wymusić ponowne odczytanie próbki z dysku.
7. Kliknij **Rozpocznij porównanie ▶**, aby uruchomić silnik dopasowywania.

---

### Krok 2: Przegląd i zatwierdzanie zmian

Po zakończeniu porównania aplikacja automatycznie przechodzi do zakładki **2. Przegląd i zatwierdzanie**.

```
+-----------------------------------------------------------------------------------------------+
| Filtr: [ Wszystkie (50) ] [ Nowe (2) ] [ Zmienione (3) ] [ Niejednoznaczne (0) ]              |
+------------------------------------+----------------------------------------------------------+
| #1 D-1001 Kowalski Jan   [NOWY]    |  IdDziecka: D-1001                                       |
| #2 D-1002 Nowak Anna     [ZMIENIONY] |  ImieNazwisko: Kowalski Jan                            |
| #3 D-1003 Wiśniewski P.  [ZMIENIONY] |  Miejscowosc: Warszawa                                 |
|                                    |  Adres: Polna 1                                          |
+------------------------------------+----------------------------------------------------------+
| [◀ Wstecz (B)] [✔ Akceptuj (A)] [✖ Odrzuć (R)] [⏭ Pomiń (S)] [✏ Zmień wartość (E)]            |
| [✔✔ Akceptuj wszystkie] [✖✖ Odrzuć wszystkie] [📊 Eksportuj raport...]   [ Zastosuj zmiany ] |
+-----------------------------------------------------------------------------------------------+
```

#### Kategorie statusów
| Status | Kolor odznaki | Opis |
|---|---|---|
| **NOWY** | Zielony (`#064E3B`) | Klucz nie występuje w bazie głównej. Wiersz zostanie dopisany na końcu bazy. |
| **ZMIENIONY** | Bursztynowy (`#713F12`) | Klucz dopasowany. Wartości różnią się w co najmniej jednej zamapowanej kolumnie. |
| **NIEJEDNOZNACZNY** | Niebieski (`#1E3A8A`) | Wiele dopasowań, zduplikowany klucz w pliku zmian lub pusty klucz. Wymaga wyboru kandydata. |
| **USUNIĘTY** | Czerwony (`#7F1D1D`) | Rekord z bazy nie występuje w pliku zmian (gdy włączona jest opcja `DetectRemovedRows`). |
| **BEZ ZMIAN** | Szary (`#1F2937`) | Klucz dopasowany; wszystkie wartości identyczne. Domyślnie ukryty; zaznacz opcję „Pokaż bez zmian” lub wybierz filtr „Bez zmian”, aby przejrzeć i sprawdzić wartości w bazie. |

#### Skróty klawiszowe i obsługa myszą
| Skrót | Akcja | Opis |
|---|---|---|
| **`A`** | **Akceptuj** | Zatwierdza proponowaną zmianę i przechodzi do następnego wiersza. |
| **`R`** | **Odrzuć** | Odrzuca proponowaną zmianę i przechodzi do następnego wiersza. |
| **`S`** | **Pomiń** | Pomija wiersz bez podejmowania decyzji. Wiersz pojawi się ponownie przy kolejnym imporcie. |
| **`E`** | **Edytuj** | Otwiera okno dialogowe umożliwiające ręczną modyfikację wartości przed zapisem. |
| **Dwuklik** | **Szybka edycja** | Podwójne kliknięcie dowolnego wiersza na liście weryfikacji natychmiast otwiera okno korekty. |
| **`B`** / **`←`** | **Wstecz** | Cofa nawigację do poprzedniego rekordu na liście. |
| **`Ctrl + Z`** | **Cofnij** | Cofa ostatnio podjętą decyzję (Undo). |
| **`F1`** | **Pomoc** | Otwiera interaktywne okno pomocy z przewodnikiem i skrótami. |

#### Selektywna aktualizacja komórek
W przypadku rekordów zmienionych w panelu szczegółów wyświetlane są wszystkie zmodyfikowane pola w postaci kart porównawczych z polem wyboru **Zastosuj**. Możesz odznaczyć pojedyncze komórki, aby zachować dotychczasową wartość z bazy, jednocześnie akceptując pozostałe zmiany.

#### Operacje zbiorcze i filtrowanie
- **Przełącznik i filtr wierszy bez zmian**: Zaznacz pole wyboru „Pokaż bez zmian” obok filtra lub wybierz opcję „Bez zmian” na liście filtrów, aby wyświetlić identyczne wiersze na liście weryfikacji bez ponownego uruchamiania porównania.
- **✔✔ Akceptuj wszystkie**: Oznacza wszystkie jednoznaczne wiersze jako zaakceptowane jednym kliknięciem (pomija wiersze bez zmian w celu ochrony istniejących danych).
- **✖✖ Odrzuć wszystkie**: Oznacza wszystkie wiersze jako odrzucone.
- **📊 Eksportuj raport**: Eksportuje całe zestawienie porównawcze do sformatowanego pliku Excel (`.xlsx`), responsywnego raportu HTML (`.html`) lub pliku tekstowego `.csv`.

---

### Krok 3: Zapis zaakceptowanych zmian do bazy

1. Gdy co najmniej jeden rekord zostanie zaakceptowany, aktywuje się przycisk **Zastosuj zaakceptowane zmiany do bazy**.
2. Kliknięcie przycisku otwiera okno potwierdzenia z podsumowaniem operacji.
3. Silnik tworzy kopię zapasową w folderze `Backups/`, dopisuje ewentualne brakujące nagłówki metadanych do wiersza 1, modyfikuje komórki w trybie InPlace i dopisuje nowe wiersze.
4. Generowane są szczegółowe dzienniki zdarzeń w `Logs/` (`.jsonl` oraz `.txt`), tworzony jest artefakt podsumowania JSON (`last_execution_summary.json` lub ścieżka `-SummaryJsonPath`), a opcjonalnie wpisy audytowe są dopisywane bezpośrednio do dedykowanego arkusza (domyślnie `ImportLog`) w pliku bazy.
5. Wyświetla się komunikat o sukcesie z liczbą zaktualizowanych komórek i potwierdzeniem zapisu arkusza logu.

---

## 4. Tryby scalania (Merge Modes) i opcje porównywania

Aplikacja oferuje zaawansowany silnik transformacji i porównywania danych, dostosowany do różnic w strukturze plików ERP, arkuszy szkolnych oraz systemów kadrowo-płacowych.

### Tryby scalania kolumn (Merge Modes)

Każda reguła w tabeli mapowania posiada konfigurowalny tryb scalania:

1. **Dokładne (Exact, 1:1)**:
   - Bezpośrednie przeniesienie wartości z jednej kolumny źródłowej do jednej kolumny bazy.
   - Gdy wartość w pliku zmian różni się od wartości w bazie, generowana jest propozycja modyfikacji komórki.

2. **Pierwsze niepuste (FirstNonEmpty, N:1)**:
   - Obsługuje logikę zastępczą (fallback coalescing).
   - Pozwala wybrać kilka potencjalnych kolumn źródłowych (np. `TelefonKomorkowy`, `TelefonDomowy`, `TelefonKontaktowy`).
   - Wartość pobierana jest z pierwszej wskazanej kolumny, która zawiera niepusty ciąg znaków.

3. **Połącz (Concatenate, N:1 oraz 1:N)**:
   - **Scalanie wielu kolumn źródłowych do jednej kolumny bazy (N:1)**:
     - Pozwala połączyć np. `UlicaINumer` + `KodPocztowy` + `Miejscowosc` w jedno pole bazy `AdresPelny` z wybranym separatorem (domyślnie `, ` lub spacja).
   - **Podział jednej kolumny źródłowej na wiele kolumn bazy (1:N)**:
     - Gdy plik zmian zawiera pełny adres w jednym polu (np. `Polna 12, 00-001 Warszawa`), a baza przechowuje go w osobnych kolumnach (`Adres` oraz `Miasto`), aplikacja inteligentnie rozdziela wartości według zadanego separatora.
   - **Dwukierunkowa eliminacja fałszywych różnic (Composite Equality)**:
     - Przed wygenerowaniem zmiany silnik weryfikuje równość kompozytową. Jeśli połączone kolumny bazy (`Adres` + `Miasto`) po złączeniu dają dokładnie taki sam ciąg jak scalona kolumna źródłowa, wiersz jest automatycznie uznawany za identyczny (`Bez zmian`).

---

### Opcje porównywania i normalizacji tekstu

W nagłówku tabeli mapowania dostępne są 4 niezależne przełączniki normalizacji:

| Opcja | Działanie | Przykład |
|---|---|---|
| **Ignoruj wielkość liter** (*Ignore Case*) | Zrównuje małe i wielkie litery podczas porównywania. | `"WARSZAWA"` == `"Warszawa"` |
| **Usuwaj spacje krawędziowe** (*Trim Whitespace*) | Usuwa zbędne spacje wiodące i końcowe przed oceną równości. | `" Jan "` == `"Jan"` |
| **Ignoruj interpunkcję i znaki specjalne** (*Ignore Punctuation*) | Pomija kropki, przecinki, myślniki, ukośniki itp. | `"ul. Polna 5/2"` == `"ul Polna 5-2"` |
| **Ignoruj białe znaki wewnątrz** (*Ignore Internal Whitespace*) | Sprowadza wielokrotne spacje, tabulatory i znaki nowej linii do pojedynczej spacji. | `"Kowalski   Jan"` == `"Kowalski Jan"` |

---

### Inteligentne reguły automatycznego mapowania (Smart Auto-Map)

Kliknięcie przycisku **⚡ Automatyczne mapowanie** uruchamia wieloetapowy algorytm heurystyczny:
1. **Dopasowanie dokładne**: Kojarzy kolumny o identycznych nazwach nagłówków (z tolerancją wielkości liter).
2. **Rozmyte dopasowanie rdzeni słów (Fuzzy Stem Matching)**:
   - Usuwa znaki interpunkcyjne i dzieli nagłówki na słowa kluczowe.
   - Automatycznie kojarzy nagłówki z literówkami (np. `pywyżej / ponieżej` ➔ `powyżej / poniżej`) oraz powszechnymi skrótami ERP (np. `niepełn.` ➔ `niepełno.`, `Adres zakładu pracy` ➔ `Adres Pracy`).
3. **Automatyczne wykrywanie scalania adresu i miasta**:
   - Gdy plik bazy posiada oddzielne kolumny dla ulicy i miasta (np. `Adres zamieszkania` oraz `Miasto Zamieszkania`), a plik zmian posiada tylko jedną połączoną kolumnę adresu, Auto-Map samoczynnie konfiguruje regułę **Połącz (Concatenate)** z separatorem `, `.
4. **Odcisk szablonu SHA-256**:
   - Po pierwszym dostosowaniu mapowania profil zapisuje się na dysku. Przy kolejnym otwarciu pliku o takich samych nagłówkach cały układ reguł wczytuje się natychmiast bez udziału użytkownika.

---

## 5. Ustawienia i kolumny metadanych

Kliknij przycisk **⚙ Ustawienia** na górnym pasku narzędzi, aby skonfigurować parametry pracy:

### Zakładka Ogólne
- **Liczba zachowywanych kopii**: Liczba archiwalnych plików kopii do zachowania (domyślnie 20). Starsze kopie są usuwane automatycznie.
- **Tryb zapisu**:
  - `InPlace` (domyślny): Szybkie patchowanie strumienia OpenXML. Zachowuje formatowanie, formuły i strukturę wieloarkuszową.
  - `SafeRewrite`: Pełne przepisanie skoroszytu za pośrednictwem FastExcelHelper.
- **Maskuj dane osobowe (PII) w logach**: Zastępuje imiona, nazwiska i adresy zamaskowanymi tokenami (`***`) w plikach dzienników oraz w arkuszu audytowym w bazie.
- **Oznacz brakujące rekordy jako usunięte**: Automatycznie wykrywa rekordy z bazy nieobecne w pliku zmian i ustawia status usunięcia.
- **Domyślnie pokazuj wiersze bez zmian**: Automatycznie włącza wyświetlanie wierszy o statusie „Bez zmian” na liście weryfikacji po każdym porównaniu.
- **Zapisuj historię zmian w arkuszu pliku bazy**: Po zaznaczeniu każda zatwierdzona modyfikacja, dopisany wiersz i usunięty rekord są rejestrowane bezpośrednio w dedykowanym arkuszu w pliku bazy Excel.
- **Nazwa arkusza logu w pliku bazy**: Własna nazwa arkusza audytowego (domyślnie: `ImportLog`, maks. 31 znaków; znaki niedozwolone `\ / ? * : [ ]` są automatycznie zamieniane na `_`).

### Zakładka Kolumny metadanych
Konfiguruje automatyczne stemple audytowe nanoszone na każdy zapisany wiersz:
| Kolumna w bazie | Domyślny token | Opis | Format |
|---|---|---|---|
| `OstZmiana` | `ChangeDate` | Data i godzina operacji zapisu | `yyyy-MM-dd HH:mm` |
| `ZrodloSciezka` | `SourceFileFullPath` | Pełna ścieżka pliku zmian | *(brak)* |
| `ZrodloPlik` | `SourceFileName` | Nazwa pliku zmian | *(brak)* |
| `Zmienil` | `CurrentUser` | Nazwa użytkownika wykonującego aktualizację | `$env:USERNAME` |
| `ZrodloWiersz` | `SourceRowNumber` | Fizyczny numer wiersza w pliku zmian | *(brak)* |
| `ImportId` | `ImportBatchId` | Unikalny identyfikator sesji GUID | *(brak)* |

---

## 6. Procedura awaryjna i przywracanie

### Przywracanie kopii z paska narzędzi
W przypadku omyłkowego zatwierdzenia niepożądanych zmian:
1. Upewnij się, że na górnym pasku wybrany jest właściwy plik bazy.
2. Kliknij przycisk **↺ Przywróć kopię** na górnym pasku.
3. Aplikacja wskaże najnowszą kopię zapasową z datą i godziną oraz zapyta o potwierdzenie.
4. Po zatwierdzeniu baza zostanie atomowo przywrócona z kopii i ponownie wczytana do widoku.

### Ochrona przed awariami zasilania i systemu
Podczas zapisu silnik posługuje się atomowymi plikami tymczasowymi (`<plik>.tmp.xlsx`).
Przy kolejnym uruchomieniu aplikacja wykrywa ewentualne osierocone pliki tymczasowe i proponuje ich usunięcie oraz przywrócenie bazy z ostatniej nienaruszonej kopii.

---

## 7. Lokalizacja kopii zapasowych i logów

### Gdzie zapisywane są kopie zapasowe?
- **Lokalizacja folderu**: Katalog `Backups/` bezpośrednio w folderze aplikacji (np. `UpdateExcelBaseFileProject\Backups\`).
- **Opcja konfiguracyjna**: Możliwość zmiany w `AppConfig.BackupDirectory` lub w oknie **⚙ Ustawienia**.
- **Szybki dostęp**: Kliknij przycisk **📁 Kopie zapasowe** na górnym pasku, aby otworzyć folder w Eksploratorze Windows.
- **Format nazwy pliku**: `<NazwaBazy>.<yyyy-MM-dd_HHmmss>.bak.xlsx` (np. `Baza_Glowna.2026-10-03_233000.bak.xlsx`).
- **Retencja i limity rozmiaru**:
  - Domyślnie przechowywanych jest ostatnich **20** kopii.
  - Automatyczne czyszczenie najstarszych kopii po przekroczeniu łącznego rozmiaru **200 MB** (`MaxBackupMb`).

### Gdzie zapisywane są pliki logów?
- **Lokalizacja folderu**: Katalog `Logs/` bezpośrednio w folderze aplikacji (np. `UpdateExcelBaseFileProject\Logs\`).
- **Opcja konfiguracyjna**: Możliwość zmiany w `AppConfig.LogDirectory` lub w oknie **⚙ Ustawienia**.
- **Szybki dostęp**: Kliknij przycisk **📋 Logi** na górnym pasku, aby otworzyć folder w Eksploratorze Windows.
- **Generowane formaty plików**:
  1. `import_<data>_<batchId>.jsonl`: Precyzyjny strumień zdarzeń w formacie JSONL rejestrujący każdą modyfikację komórki.
  2. `import_<data>_<batchId>.txt`: Czytelny raport podsumowujący ze statystykami i tabelarycznym wykazem wierszy.
  3. `masterlog.jsonl`: Zbiorcza historia wszystkich operacji w aplikacji.
- **Ochrona danych osobowych (PII)**: Przy włączonej opcji maskowania nazwiska, numery telefonów i adresy są zastępowane tokenami ochronnymi.

---

## 8. Podpowiedzi, skróty i pomoc w aplikacji

- **Interaktywna pomoc w aplikacji**: Wciśnij klawisz **`F1`** w dowolnym momencie lub kliknij przycisk **❓ Pomoc** na górnym pasku narzędzi, aby otworzyć okno pomocy z kartami objaśniającymi procedurę, tryby scalania, opcje porównywania, lokalizację plików i listę skrótów.
- **Opisowe dymki pomocy (tooltips)**: Najedź kursorem myszy na dowolny przycisk, pole edycyjne lub nagłówek, aby zobaczyć szczegółowe objaśnienie działania danej funkcji.
- **Wsparcie wielojęzyczne**: Cały interfejs, komunikaty, dymki pomocy oraz dokumentacja dostępne są w językach: **polskim (Polski)**, **angielskim (English)** oraz **niemieckim (Deutsch)**.

---

# Excel Master Updater — Deutsches Benutzerhandbuch

## 1. Übersicht und Zweck

**Excel Master Updater** ist eine Windows-Desktop-Anwendung zur zuverlässigen Pflege und Synchronisierung einer zentralen **Master-Excel-Basisdatei** anhand wiederkehrender Änderungs- und Aktualisierungsdateien.

### Wichtigste Funktionen
- **Universelle Spaltenzuordnung**: Ordnen Sie Kopfzeilen einmalig zu; nachfolgende Dateien mit identischer Tabellenstruktur werden über einen SHA-256-Fingerabdruck automatisch erkannt.
- **Zeilenweise visuelle Prüfung**: Überprüfen Sie jeden Änderungsvorschlag auf übersichtlichen Differenzkarten mit individuellen Kontrollkästchen je Zelle und schneller Doppelklick-Bearbeitung.
- **Sichere InPlace OpenXML-Engine**: Schreibt Änderungen direkt in den XML-Stream des Arbeitsblatts, ohne Zellenfarben, Schriftarten, Spaltenbreiten oder andere Blätter zu verändern.
- **Robuste CSV-Engine**: Automatische Erkennung von Trennzeichen (Semikolon `;`, Komma `,`, Tabulator `\t`, Pipe `|`) sowie Zeichencodierungen (UTF-8 mit/ohne BOM, UTF-16 und Windows-1250 ANSI-Fallback).
- **Automatische Revisionsstempel**: Hinterlegt Änderungsdatum, Quelldateiname, Benutzername, Quellzeilennummer und Batch-GUID direkt in den Metadatenspalten der Basisdatei.
- **Automatische Sicherungen und JSON-Zusammenfassungen**: Erstellt vor jedem Schreibvorgang eine zeitgestempelte Sicherungskopie im Ordner `Backups/` und exportiert maschinenlesbare JSON-Zusammenfassungen für CI/CD-Pipelines.

---

## 2. Schnellstart

### Starten der Anwendung
Start über PowerShell 7 (empfohlen):
```powershell
pwsh -File .\start.ps1
```
Oder über Windows PowerShell 5.1:
```powershell
powershell -File .\start.ps1
```
Oder direkt über die kompilierte Binärdatei:
```powershell
.\Master-Updater.exe
```

### Vorab-Laden von Dateien über die Befehlszeile
```powershell
pwsh -File .\start.ps1 -BaseFilePath "C:\Daten\Basis_Master.xlsx" -IncomingPath "C:\Daten\Aenderungen_September.xlsx"
```

### Unbeaufsichtigte Batch-Ausführung (Headless)
Führen Sie einen vollautomatischen Abgleich ohne grafische Oberfläche aus:
```powershell
pwsh -File .\Master-Updater.ps1 -BaseFilePath "C:\Daten\Basis_Master.xlsx" -IncomingPath "C:\Daten\Aenderungen.csv" -Headless -AutoAccept AllNonAmbiguous -ExportReportPath "C:\Berichte\diff.html" -SummaryJsonPath "C:\Logs\summary.json"
```

---

## 3. Arbeitsablauf

### Schritt 1: Dateiauswahl und Spaltenzuordnung
1. **Basisdatei auswählen**: Wählen Sie Ihre Master-Excel-Datei (`.xlsx`) oder `.csv`-Datei aus. Wählen Sie das gewünschte Datenblatt aus.
2. **Änderungsdatei auswählen**: Ziehen Sie die Datei per Drag-and-Drop hinein oder wählen Sie sie über den Dateidialog aus (`.xlsx` oder `.csv`). Trennzeichen (`;`, `,`, `\t`, `|`) und Codierungen werden automatisch erkannt. Wählen Sie das Quellarbeitsblatt.
3. **Automatische Vorlagenerkennung**:
   - Wurden die Kopfzeilen bereits zuvor zugeordnet, erkennt die Anwendung das Vorlagen-Profil anhand des SHA-256-Fingerabdrucks automatisch.
   - Handelt es sich um eine neue Vorlage, klicken Sie auf **⚡ Auto-Zuordnung** oder **+ Regel hinzufügen**, um Spalten zu verknüpfen.
4. **Zuordnungsregeln filtern und durchsuchen**:
   - Geben Sie in der Symbolleiste der Zuordnungstabelle einen Suchbegriff in das Feld **Zuordnungsregeln filtern...** (`txtSearchMapping`) ein, um Ziel- oder Quellspalten sofort einzugrenzen. Klicken Sie auf **`✖`**, um den Filter zu löschen.
5. **Verknüpfungsschlüssel und Live-Indikator**:
   - Markieren Sie in den Schlüssellisten die Spalten zur eindeutigen Identifikation (z. B. `Id` in der Basis und `SchuelerId` in der Änderungsdatei, oder zusammengesetzte Schlüssel wie `Nachname` + `Vorname`).
   - Jede Änderung der Schlüsselauswahl aktualisiert sofort die **Daten- und Zuordnungsvorschau** unten und markiert die Schlüssel in der Zielspalte mit einem Schlüsselsymbol (**`🔑 `**).
6. **Echtzeit-Daten- und Zuordnungsvorschau (Unterer Bereich)**:
   - Überprüfen Sie Beispieldatenzeilen und testen Sie Ihre Zuordnungsregeln in Echtzeit am unteren Bildschirmrand vor dem Start des Abgleichs:
     - **⚡ Zuordnungsergebnis (Vorschau)**: Zeigt alle Basisspalten, die zugeordneten Quellspalten, den projizierten Wert, den aktuellen Basiswert und ein sofortiges Statusabzeichen (`Identisch`, `Änderung`, `Neuer Wert` oder `Nicht zugeordnet`).
     - **📥 Änderungsdatei (Musterzeile)**: Zeigt eine echte horizontale Musterzeile aus der Änderungsdatei an.
     - **🗄️ Basisdatei (Musterzeile)**: Zeigt eine echte horizontale Musterzeile aus der Master-Basisdatei an.
   - Verwenden Sie **`◀ Vorherige Zeile`** und **`Nächste Zeile ▶`**, um durch die Musterzeilen 1 bis 10 zu blättern.
   - Klicken Sie auf **`🔄 Vorschau aktualisieren`**, um ein erneutes Einlesen von der Festplatte zu erzwingen.
7. Klicken Sie auf **Vergleich starten ▶**, um den Vergleich auszuführen.

---

### Schritt 2: Prüfung und Freigabe

Nach Abschluss des Vergleichs wechselt die Anwendung automatisch auf die Registerkarte **2. Prüfung & Freigabe**.

```
+-----------------------------------------------------------------------------------------------+
| Filter: [ Alle (50) ] [ Neu (2) ] [ Geändert (3) ] [ Mehrdeutig (0) ]                         |
+------------------------------------+----------------------------------------------------------+
| #1 D-1001 Schmidt Hans   [NEU]     |  SchuelerId: D-1001                                      |
| #2 D-1002 Weber Anna     [GEÄNDERT]|  Name: Schmidt Hans                                      |
| #3 D-1003 Braun Peter    [GEÄNDERT]|  Stadt: Berlin                                           |
|                                    |  Strasse: Hauptstrasse 1                                 |
+------------------------------------+----------------------------------------------------------+
| [◀ Zurück (B)] [✔ Akzeptieren (A)] [✖ Ablehnen (R)] [⏭ Überspringen (S)] [✏ Wert bearbeiten]    |
| [✔✔ Alle akzeptieren] [✖✖ Alle ablehnen] [📊 Bericht exportieren...]  [ Änderungen anwenden ]  |
+-----------------------------------------------------------------------------------------------+
```

#### Status-Kategorien
| Status | Abzeichen-Farbe | Beschreibung |
|---|---|---|
| **NEU** | Grün (`#064E3B`) | Schlüssel existiert nicht in der Basisdatei. Zeile wird am Ende angehängt. |
| **GEÄNDERT** | Bernstein (`#713F12`) | Schlüssel übereinstimmend. Werte weichen in mindestens einer zugeordneten Spalte ab. |
| **MEHRDEUTIG** | Blau (`#1E3A8A`) | Mehrere Treffer, doppelter Schlüssel in Quelldaten oder leerer Schlüssel. Erfordert manuelle Auswahl. |
| **ENTFERNT** | Rot (`#7F1D1D`) | Basisdatensatz fehlt in der Änderungsdatei (wenn `DetectRemovedRows` aktiviert ist). |
| **UNVERÄNDERT** | Grau (`#1F2937`) | Schlüssel übereinstimmend; alle Werte identisch. Standardmäßig ausgeblendet; aktivieren Sie „Unveränderte anzeigen“ oder wählen Sie den Filter „Unverändert“, um Basiswerte einzusehen und zu prüfen. |

#### Tastatur- und Maus-Kürzel
| Tastenkürzel | Aktion | Beschreibung |
|---|---|---|
| **`A`** | **Akzeptieren** | Übernimmt die Zeile als freigegeben und springt zum nächsten Datensatz. |
| **`R`** | **Ablehnen** | Verwirft die Zeile und springt zum nächsten Datensatz. |
| **`S`** | **Überspringen** | Belässt den Datensatz ohne Entscheidung. Erscheint beim nächsten Import erneut. |
| **`E`** | **Bearbeiten** | Öffnet ein Dialogfenster zur manuellen Korrektur vorgeschlagener Werte vor dem Schreiben. |
| **Doppelklick** | **Schnelle Bearbeitung** | Ein Doppelklick auf einen Datensatz in der Liste öffnet sofort das Bearbeitungsfenster. |
| **`B`** / **`←`** | **Zurück** | Navigiert zum vorherigen Datensatz zurück. |
| **`Strg + Z`** | **Rückgängig** | Macht die letzte Entscheidung rückgängig (Undo). |
| **`F1`** | **Hilfe** | Öffnet das interaktive Hilfefenster mit Kurzanleitung und Verzeichnissen. |

#### Selektive Zellenaktualisierung
Bei geänderten Datensätzen zeigt der Detailbereich alle abweichenden Felder nebeneinander auf Differenzkarten mit einem Kontrollkästchen **Übernehmen** an. Sie können einzelne Zellen abwählen, um den bestehenden Basiswert beizubehalten, während andere Änderungen übernommen werden.

#### Massenwerkzeuge und Filterung
- **Umschalter und Filter für unveränderte Zeilen**: Aktivieren Sie das Kontrollkästchen „Unveränderte anzeigen“ neben der Filterauswahl oder wählen Sie „Unverändert“ im Statusfilter, um identische Zeilen ohne erneuten Vergleichslauf in der Prüfliste anzuzeigen.
- **✔✔ Alle akzeptieren**: Markiert alle eindeutigen Zeilen mit einem Klick als akzeptiert (überspringt unveränderte Zeilen zum Schutz bestehender Werte).
- **✖✖ Alle ablehnen**: Markiert alle Zeilen als abgelehnt.
- **📊 Bericht exportieren**: Exportiert die gesamte Prüfungsübersicht nach Microsoft Excel (`.xlsx`), als eigenständigen responsiven HTML-Bericht (`.html`) oder als `.csv`.

---

### Schritt 3: Zurückschreiben in die Basisdatei

1. Sobald mindestens ein Datensatz akzeptiert wurde, wird die Schaltfläche **Akzeptierte Änderungen anwenden** aktiv.
2. Klicken Sie auf die Schaltfläche, um den Bestätigungsdialog zu öffnen.
3. Die Engine erstellt eine automatische Sicherungskopie im Ordner `Backups/`, ergänzt fehlende Metadaten-Kopfzeilen in Zeile 1, modifiziert Zellen direkt (InPlace) und hängt neue Zeilen an.
4. Es werden detaillierte Protokolldateien in `Logs/` (`.jsonl` und `.txt`) erstellt, ein maschinenlesbares Ausführungs-JSON generiert (`last_execution_summary.json` oder über `-SummaryJsonPath`) und optional alle Änderungen direkt in ein dediziertes Protokollblatt (Standard: `ImportLog`) in der Basisdatei geschrieben.
5. Es wird eine Erfolgsmeldung mit der Anzahl aktualisierter Zellen und der Bestätigung des Protokollblatts angezeigt.

---

## 4. Zusammenführungsmodi (Merge Modes) und Vergleichsoptionen

Die Anwendung bietet eine leistungsfähige Transformations- und Normalisierungs-Engine zum zuverlässigen Abgleich abweichender Datenstrukturen aus ERP-Exporten, Schulverwaltungssystemen und Personaldateien.

### Spalten-Zusammenführungsmodi (Merge Modes)

Jede Zuordnungsregel unterstützt drei konfigurierbare Zusammenführungsmodi:

1. **Exakt (Exact, 1:1)**:
   - Direkte 1-zu-1-Übertragung aus einer Quellspalte in eine Basisspalte.
   - Löst einen Änderungsvorschlag aus, sobald der Quellwert vom aktuellen Basiswert abweicht.

2. **Erstes nicht-leeres (FirstNonEmpty, N:1)**:
   - Realisiert eine Ausweichlogik (Fallback-Coalescing) über mehrere Quellspalten.
   - Ermöglicht die Auswahl mehrerer Spalten (z. B. `Mobiltelefon`, `Festnetz`, `Diensttelefon`).
   - Übernimmt den Wert der ersten Quellspalte, die nicht leer ist.

3. **Zusammenführen (Concatenate, N:1 und 1:N)**:
   - **N:1-Zusammenführung (Mehrere Quellspalten in eine Basisspalte)**:
     - Verbindet Felder wie `StrasseHausnummer` + `Postleitzahl` + `Ort` zu einer Basisspalte `VollstaendigeAdresse` mit frei wählbarem Trennzeichen (z. B. `, ` oder Leerzeichen).
   - **1:N-Aufteilung (Eine Quellspalte auf mehrere Basisspalten)**:
     - Liefert die Quelldatei eine kombinierte Adresse (z. B. `Musterstrasse 1, 10115 Berlin`), die Basisdatei jedoch separate Spalten (`Adresse` und `Stadt`), teilt das System den Wert anhand des Trennzeichens auf.
   - **Zusammengesetzte Gleichheitsprüfung (Composite Equality)**:
     - Vor dem Auslösen von Differenzen prüft die Engine, ob die kombinierten Basisspalten dem zusammengeführten Quelltext entsprechen. Ist dies der Fall, wird der Datensatz als `Unverändert` eingestuft (keine Scheinunterschiede).

---

### Vergleichs- und Textnormalisierungsoptionen

In der Kopfzeile der Zuordnungstabelle stehen 4 Normalisierungsoptionen zur Verfügung:

| Option | Wirkung | Beispiel |
|---|---|---|
| **Groß-/Kleinschreibung ignorieren** (*Ignore Case*) | Behandelt Groß- und Kleinbuchstaben als identisch. | `"BERLIN"` == `"Berlin"` |
| **Randleerzeichen kürzen** (*Trim Whitespace*) | Entfernt führende und nachgestellte Leerzeichen vor dem Vergleich. | `" Hans "` == `"Hans"` |
| **Satzzeichen & Sonderzeichen ignorieren** (*Ignore Punctuation*) | Vergleicht alphanumerische Zeichen unter Auslassung von Punkten, Kommas, Bindestrichen usw. | `"Hauptstr. 5/2"` == `"Hauptstr 5-2"` |
| **Interne Leerzeichen ignorieren** (*Ignore Internal Whitespace*) | Fasst mehrfache Leerzeichen, Tabulatoren und Zeilenumbrüche zu einem Einzelleerzeichen zusammen. | `"Müller   Hans"` == `"Müller Hans"` |

---

### Intelligente Auto-Zuordnungsregeln (Smart Auto-Map)

Ein Klick auf **⚡ Auto-Zuordnung** aktiviert einen mehrstufigen Heuristik-Algorithmus:
1. **Exakte und schreibungsunabhängige Zuordnung**: Ordnet Spalten mit identischen Namen direkt zu.
2. **Unscharfe Wortstamm-Erkennung (Fuzzy Stem Matching)**:
   - Entfernt Satzzeichen und vergleicht Wortstämme (Mindestüberdeckung 50 %).
   - Toleriert gängige Tippfehler (z. B. `pywyżej` ➔ `powyżej`) sowie ERP-Abkürzungen (z. B. `niepełn.` ➔ `niepełno.`, `Adres zakładu pracy` ➔ `Adres Pracy`).
3. **Automatische Erkennung kombinierter Adress- und Stadtspalten**:
   - Sind in der Basisdatei Adresse und Stadt getrennt, in der Änderungsdatei jedoch kombiniert, richtet die Auto-Zuordnung automatisch eine **Zusammenführen (Concatenate)**-Regel mit Trennzeichen `, ` ein.
4. **SHA-256-Vorlagen-Fingerabdruck**:
   - Die Zuordnung wird gespeichert und bei zukünftigen Importen identischer Kopfzeilenlayouts automatisch wiederhergestellt.

---

## 5. Einstellungen und Metadatenspalten

Klicken Sie in der oberen Symbolleiste auf **⚙ Einstellungen**, um Optionen anzupassen:

### Registerkarte Allgemein
- **Anzahl Sicherungen**: Anzahl aufzubewahrender historischer Sicherungen (Standard: 20). Ältere Sicherungen werden automatisch gelöscht.
- **Schreibmodus**:
  - `InPlace` (Standard): Schnelles OpenXML-Stream-Patching. Behält Formatierungen, Stile und mehrere Arbeitsblätter vollständig bei.
  - `SafeRewrite`: Vollständige Neuerstellung der Arbeitsmappe über FastExcelHelper.
- **Maskierung personenbezogener Daten (PII)**: Ersetzt Namen und Adressen in Protokolldateien sowie im Protokollblatt durch Platzhalter (`***`).
- **Fehlende Datensätze als entfernt markieren**: Erkennt in der Basisdatei vorhandene, aber in der Änderungsdatei fehlende Zeilen automatisch.
- **Unveränderte Zeilen standardmäßig anzeigen**: Aktiviert die Anzeige von Zeilen mit dem Status „Unverändert“ in der Prüfliste standardmäßig nach jedem Vergleich.
- **Änderungen in einem Tabellenblatt der Basisdatei protokollieren**: Protokolliert jede Zellenänderung, hinzugefügte Zeile und entfernte Datensätze direkt in einem Arbeitsblatt der Basis-Excel-Datei.
- **Name des Protokollblatts in der Basisdatei**: Benutzerdefinierter Name des Revisionsblatts (Standard: `ImportLog`, max. 31 Zeichen; unzulässige Zeichen `\ / ? * : [ ]` werden automatisch durch `_` ersetzt).

### Registerkarte Metadatenspalten
Legt die automatischen Revisionsstempel fest, die in jede aktualisierte Zeile geschrieben werden:
| Basis-Spalte | Standard-Token | Beschreibung | Format |
|---|---|---|---|
| `OstZmiana` | `ChangeDate` | Datum und Uhrzeit der Aktualisierung | `yyyy-MM-dd HH:mm` |
| `ZrodloSciezka` | `SourceFileFullPath` | Vollständiger Pfad der Änderungsdatei | *(kein)* |
| `ZrodloPlik` | `SourceFileName` | Dateiname der Änderungsdatei | *(kein)* |
| `Zmienil` | `CurrentUser` | Anmeldename des ausführenden Benutzers | `$env:USERNAME` |
| `ZrodloWiersz` | `SourceRowNumber` | Physische Zeilennummer in der Quelldatei | *(kein)* |
| `ImportId` | `ImportBatchId` | Eindeutige Sitzungs-GUID | *(kein)* |

---

## 6. Notfallwiederherstellung & Restore

### Wiederherstellung über die Symbolleiste
Sollte versehentlich ein unerwünschter Schreibvorgang bestätigt worden sein:
1. Vergewissern Sie sich, dass der Pfad der Basisdatei in der oberen Leiste ausgewählt ist.
2. Klicken Sie in der oberen Symbolleiste auf **↺ Sicherung wiederherstellen**.
3. Die Anwendung zeigt die jüngste Sicherung mit Zeitstempel an und bittet um Bestätigung.
4. Nach Bestätigung wird die Basisdatei atomar aus der Sicherung wiederhergestellt und in der Oberfläche neu geladen.

### Absturzsicherung
Während des Schreibens verwendet die Engine temporäre Staging-Dateien (`<Datei>.tmp.xlsx`).
Beim nächsten Start erkennt die Anwendung verwaiste Zwischendateien automatisch und bietet an, diese zu bereinigen und die letzte intakte Sicherung zu laden.

---

## 7. Speicherorte für Sicherungen und Protokolle

### Wo werden Sicherungskopien gespeichert?
- **Ordnerpfad**: Ordner `Backups/` direkt im Installationsverzeichnis der Anwendung (z. B. `UpdateExcelBaseFileProject\Backups\`).
- **Konfigurationsoption**: Anpassbar über `AppConfig.BackupDirectory` oder in den **⚙ Einstellungen**.
- **Schnellzugriff**: Klicken Sie in der Symbolleiste auf **📁 Sicherungen**, um den Ordner direkt im Windows-Explorer zu öffnen.
- **Namensformat**: `<BasisDateiname>.<yyyy-MM-dd_HHmmss>.bak.xlsx` (z. B. `Basis_Master.2026-10-03_233000.bak.xlsx`).
- **Aufbewahrung & Speicherbegrenzung**:
  - Standardmäßig werden die letzten **20** Sicherungen aufbewahrt.
  - Automatisches Löschen ältester Sicherungen, sobald die Gesamtgröße **200 MB** überschreitet (`MaxBackupMb`).

### Wo werden Protokolldateien gespeichert?
- **Ordnerpfad**: Ordner `Logs/` direkt im Installationsverzeichnis der Anwendung (z. B. `UpdateExcelBaseFileProject\Logs\`).
- **Konfigurationsoption**: Anpassbar über `AppConfig.LogDirectory` oder in den **⚙ Einstellungen**.
- **Schnellzugriff**: Klicken Sie in der Symbolleiste auf **📋 Protokolle**, um den Ordner im Windows-Explorer zu öffnen.
- **Erstellte Dateiformate**:
  1. `import_<Datum>_<BatchId>.jsonl`: Detaillierter Ereignisdatenstrom im JSON-Lines-Format für jede Zellenmodifikation.
  2. `import_<Datum>_<BatchId>.txt`: Übersichtlicher, lesbarer Textbericht mit Kennzahlen und tabellarischer Zeilenübersicht.
  3. `masterlog.jsonl`: Fortlaufende Gesamthistorie aller ausgeführten Synchronisierungsläufe.
- **Datenschutz (PII)**: Ist die Maskierung aktiviert, werden Personennamen, Telefonnummern und Adressen in den Protokollen unkenntlich gemacht.

---

## 8. QuickInfo-Texte, Tastaturkürzel und In-App-Hilfe

- **Interaktive Hilfe in der Anwendung**: Drücken Sie jederzeit **`F1`** oder klicken Sie in der Symbolleiste auf **❓ Hilfe**, um das Hilfefenster mit Workflow-Karten, Zusammenführungsmodi, Vergleichsoptionen, Verzeichnis-Schaltflächen und Tastaturkürzeln zu öffnen.
- **Detaillierte QuickInfo-Texte (Tooltips)**: Bewegen Sie die Maus über eine beliebige Schaltfläche oder ein Eingabefeld, um eine präzise Erklärung der Funktionsweise und zugehörige Tastenkürzel einzublenden.
- **Dreisprachigkeit**: Die gesamte Oberfläche, Dialoge, Meldungen und Handbücher sind vollständig auf **Deutsch**, **Englisch** und **Polnisch** verfügbar.
