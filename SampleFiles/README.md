# Sample Files for Excel Master Updater

This folder contains realistic, production-grade test and demonstration Excel (`.xlsx`) files for school transport management (*dowóz uczniów niepełnosprawnych i szkolnych*).

Both files were generated using pure .NET OpenXML compression (`Generate-SampleFiles.ps1`) without external dependencies or COM automation.

---

## Files Overview

| File | Purpose | Columns | Rows | Sheet Name |
| :--- | :--- | :--- | :--- | :--- |
| **`Baza_Uczniowie_Dowoz.xlsx`** | Master base database | 15 | 20 | `Dowóz Uczniów` |
| **`Zmiany_Wrzesien_2026.xlsx`** | Incoming monthly change batch | 10 | 16 | `Wnioski Wrzesień` |
| **`Generate-SampleFiles.ps1`** | Deterministic generator script | - | - | - |
| **`Raport_Roznic_Probka.html`** | Standalone dark-mode HTML diff report | - | - | - |

---

## Column Schemas & Mapping

### Master Base File (`Baza_Uczniowie_Dowoz.xlsx`)

| Col # | Header Name | Type / Notes |
| :---: | :--- | :--- |
| **A** | `Imię i Nazwisko Dziecka` | Primary join key (student full name) |
| **B** | `Szkoła` | Target educational institution |
| **C** | `Adres` | Home address of the student (base-only) |
| **D** | `Rodzaj dowozu` | Transport type (e.g. bus line or own transport) |
| **E** | `Opiekun` | Legal guardian name |
| **F** | `Telefon` | Guardian contact number |
| **G** | `Łączenie dowozu z pracą` | Commute combining flag (`Tak` / `Nie`) |
| **H** | `Adres Praca` | Guardian workplace address (base-only) |
| **I** | `Samochód` | Vehicle make |
| **J** | `Model` | Vehicle model |
| **K** | `Pojemność` | Engine displacement (cm³) |
| **L** | `Uwagi` | Special remarks / assistant requirements |
| **M** | `Ostatnia Zmiana` | Audit column: timestamp of last update |
| **N** | `Plik zmiany` | Audit column: incoming batch filename |
| **O** | `Aktualny` | Status flag (`Tak` / `Nie`) |

### Incoming Changes File (`Zmiany_Wrzesien_2026.xlsx`)

| Col # | Header Name | Maps To Base Column |
| :---: | :--- | :--- |
| **A** | `Imię i Nazwisko Dziecka` | `Imię i Nazwisko Dziecka` (Join Key) |
| **B** | `Szkoła` | `Szkoła` |
| **C** | `Rodzaj dowozu` | `Rodzaj dowozu` |
| **D** | `Opiekun` | `Opiekun` |
| **E** | `Telefon` | `Telefon` |
| **F** | `Łączenie dowozu z pracą` | `Łączenie dowozu z pracą` |
| **G** | `Samochód` | `Samochód` |
| **H** | `Model` | `Model` |
| **I** | `Pojemność` | `Pojemność` |
| **J** | `Uwagi` | `Uwagi` |

---

## Comparison Test Scenarios

The sample incoming file includes 16 records covering all comparison scenarios:

1. **Unchanged Records (10)**:
   - Records 1-3, 5, 7-10, 12-13: exact match with base file data.
2. **Changed Records (4)**:
   - `Wójcik Michał`: Transferred from Gminny Bus to own transport; new phone number; vehicle details added.
   - `Lewandowski Aleksander`: Vehicle switched from Opel Astra to Dacia Jogger; updated psychological decision in remarks.
   - `Kozłowska Alicja`: Phone number updated; commute combination changed.
   - `Krawczyk Laura`: Transferred to a new school (`Zespół Szkół Specjalnych nr 2`).
3. **New Records (2)**:
   - `Borkowski Stanisław`: Newly enrolled student.
   - `Czarnecka Helena`: Newly enrolled student.
4. **Base-Only Columns Preserved**:
   - `Adres`, `Adres Praca`, `Ostatnia Zmiana`, `Plik zmiany`, and `Aktualny` are preserved without being overwritten.

---

## Quick Start Commands

### 1. Run Headless Batch Review & Generate HTML Diff Report

```powershell
pwsh -NoProfile -Command "
  . .\Master-Updater.ps1
  Invoke-HeadlessMasterUpdater `
    -BaseFilePath '.\SampleFiles\Baza_Uczniowie_Dowoz.xlsx' `
    -IncomingPath '.\SampleFiles\Zmiany_Wrzesien_2026.xlsx' `
    -ExportReportPath '.\SampleFiles\Raport_Roznic_Probka.html'
"
```

### 2. Launch GUI and Load Sample Files

```powershell
pwsh -NoProfile -File .\start.ps1
```

In the GUI:
1. Under **Plik Bazy (Master)**, browse to `SampleFiles\Baza_Uczniowie_Dowoz.xlsx`.
2. Under **Plik Wejściowy**, browse to `SampleFiles\Zmiany_Wrzesien_2026.xlsx`.
3. Verify matching columns in the **Mapowanie** tab.
4. Click **Porównaj i Przejrzyj (F5)** to see real-time color diffs and approve changes.

### 3. Regenerate Sample Files

To reset the sample files back to their pristine baseline:

```powershell
pwsh -NoProfile -File .\SampleFiles\Generate-SampleFiles.ps1
```
