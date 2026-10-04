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

## 4. Einstellungen und Metadatenspalten

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

## 5. Notfallwiederherstellung & Restore

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

## 6. Speicherorte für Sicherungen und Protokolle

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

## 7. QuickInfo-Texte, Tastaturkürzel und In-App-Hilfe

- **Interaktive Hilfe in der Anwendung**: Drücken Sie jederzeit **`F1`** oder klicken Sie in der Symbolleiste auf **❓ Hilfe**, um das Hilfefenster mit Workflow-Karten, Verzeichnis-Schaltflächen und Tastaturkürzeln zu öffnen.
- **Detaillierte QuickInfo-Texte (Tooltips)**: Bewegen Sie die Maus über eine beliebige Schaltfläche oder ein Eingabefeld, um eine präzise Erklärung der Funktionsweise und zugehörige Tastenkürzel einzublenden.
- **Dreisprachigkeit**: Die gesamte Oberfläche, Dialoge, Meldungen und Handbücher sind vollständig auf **Deutsch**, **Englisch** und **Polnisch** verfügbar.
