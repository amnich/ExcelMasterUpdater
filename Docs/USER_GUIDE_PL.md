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
