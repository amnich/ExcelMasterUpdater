<#
.SYNOPSIS
    Generates deterministic test fixtures for Excel Master Updater regression testing.

.DESCRIPTION
    Creates the complete set of synthetic test fixture files without external dependencies:
    - base.xlsx: 50 rows master file with 20 data columns + 6 metadata columns.
    - incoming_A.xlsx: 50 rows incoming source with address split, 3 changed rows, 2 new rows, 45 unchanged rows.
    - incoming_B.csv: Semicolon-delimited CSV test file containing empty join key and duplicate key edge cases.
    - base_formatted.xlsx: Multi-sheet workbook with custom cell styling to test InPlace formatting preservation.

.PARAMETER OutputDir
    Target directory where generated test fixture files will be saved. Defaults to $PSScriptRoot.

.EXAMPLE
    pwsh -File .\Generate-TestFixtures.ps1
    Generates all fixture files in the current directory.

.OUTPUTS
    System.Void. Generates .xlsx and .csv files on disk.

.NOTES
    Pure .NET System.IO.Compression and System.Xml. Zero Excel COM or third-party dependencies.
#>
param(
    [string]$OutputDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml

# Helper to escape XML
function Escape-XmlText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}

# Helper to convert column 0-based index to Excel name (0 -> A, 27 -> AB)
function ConvertTo-ExcelColName([int]$index) {
    $dividend = $index + 1
    $colName = ''
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $colName = [char](65 + $modulo) + $colName
        $dividend = [int]([Math]::Floor(($dividend - $modulo) / 26))
    }
    return $colName
}

# Generates a minimal OpenXML .xlsx file from headers and row hashtables
function New-SimpleExcelFile {
    param(
        [string]$FilePath,
        [string]$SheetName = 'Sheet1',
        [string[]]$Headers,
        [System.Collections.IList]$Rows,
        [string]$SecondSheetName = $null,
        [System.Collections.IList]$SecondSheetRows = $null
    )

    if (Test-Path $FilePath) { Remove-Item -Force $FilePath }
    $dir = Split-Path $FilePath -Parent
    if (-not (Test-Path $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }

    $tempZip = [System.IO.Path]::GetTempFileName()
    if (Test-Path $tempZip) { Remove-Item -Force $tempZip }

    $fs = $null
    $zip = $null
    try {
        $fs = [System.IO.File]::Create($tempZip)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

        # [Content_Types].xml
        $hasSecond = -not [string]::IsNullOrEmpty($SecondSheetName)
        $ctEntry = $zip.CreateEntry('[Content_Types].xml')
        $sw = New-Object System.IO.StreamWriter($ctEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $secondCt = if ($hasSecond) { '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' } else { '' }
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/xl/workbook.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`"/><Override PartName=`"/xl/worksheets/sheet1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>$secondCt<Override PartName=`"/xl/styles.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`"/></Types>")
        } finally { $sw.Dispose() }

        # _rels/.rels
        $relsEntry = $zip.CreateEntry('_rels/.rels')
        $sw = New-Object System.IO.StreamWriter($relsEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>")
        } finally { $sw.Dispose() }

        # xl/_rels/workbook.xml.rels
        $wbRelsEntry = $zip.CreateEntry('xl/_rels/workbook.xml.rels')
        $sw = New-Object System.IO.StreamWriter($wbRelsEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $secondRel = if ($hasSecond) { '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' } else { '' }
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet1.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>$secondRel</Relationships>")
        } finally { $sw.Dispose() }

        # xl/workbook.xml
        $wbEntry = $zip.CreateEntry('xl/workbook.xml')
        $sw = New-Object System.IO.StreamWriter($wbEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $s1Esc = Escape-XmlText $SheetName
            $secondSheetXml = if ($hasSecond) { "<sheet name=`"$(Escape-XmlText $SecondSheetName)`" sheetId=`"2`" r:id=`"rId3`"/>" } else { '' }
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets><sheet name=`"$s1Esc`" sheetId=`"1`" r:id=`"rId1`"/>$secondSheetXml</sheets></workbook>")
        } finally { $sw.Dispose() }

        # xl/styles.xml with formatted styles
        $stylesEntry = $zip.CreateEntry('xl/styles.xml')
        $sw = New-Object System.IO.StreamWriter($stylesEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><fonts count=`"3`"><font><sz val=`"11`"/><name val=`"Segoe UI`"/></font><font><b/><sz val=`"11`"/><color rgb=`"FFFFFFFF`"/><name val=`"Segoe UI`"/></font><font><i/><sz val=`"10`"/><name val=`"Segoe UI`"/></font></fonts><fills count=`"4`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF1E3A8A`"/></fill></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"FFFEF3C7`"/></fill></fill></fills><borders count=`"2`"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=`"thin`"><color rgb=`"FFCBD5E1`"/></left><right style=`"thin`"><color rgb=`"FFCBD5E1`"/></right><top style=`"thin`"><color rgb=`"FFCBD5E1`"/></top><bottom style=`"thin`"><color rgb=`"FFCBD5E1`"/></bottom><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"3`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"1`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"2`" borderId=`"1`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`"/><xf numFmtId=`"0`" fontId=`"0`" fillId=`"3`" borderId=`"1`" xfId=`"0`" applyFill=`"1`" applyBorder=`"1`"/></cellXfs></styleSheet>")
        } finally { $sw.Dispose() }

        # xl/worksheets/sheet1.xml
        $s1Entry = $zip.CreateEntry('xl/worksheets/sheet1.xml')
        $sw = New-Object System.IO.StreamWriter($s1Entry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $totalRows = $Rows.Count + 1
            $lastColName = ConvertTo-ExcelColName ($Headers.Count - 1)
            $dimRef = "A1:$($lastColName)$($totalRows)"

            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><dimension ref=`"$dimRef`"/><sheetViews><sheetView tabSelected=`"1`" workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=`"15`"/><sheetData>")

            # Header row (row 1, style s=1)
            $sw.Write("<row r=`"1`" spans=`"1:$($Headers.Count)`">")
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $colLetter = ConvertTo-ExcelColName $c
                $hText = Escape-XmlText $Headers[$c]
                $sw.Write("<c r=`"$($colLetter)1`" s=`"1`" t=`"inlineStr`"><is><t>$hText</t></is></c>")
            }
            $sw.Write("</row>")

            # Data rows
            for ($r = 0; $r -lt $Rows.Count; $r++) {
                $rowNum = $r + 2
                $rowData = $Rows[$r]
                $sw.Write("<row r=`"$rowNum`" spans=`"1:$($Headers.Count)`">")
                for ($c = 0; $c -lt $Headers.Count; $c++) {
                    $hName = $Headers[$c]
                    $val = if ($rowData -is [System.Collections.IDictionary] -and $rowData.Contains($hName)) {
                        $rowData[$hName]
                    } elseif ($rowData.PSObject.Properties[$hName]) {
                        $rowData.$hName
                    } else { '' }

                    $valStr = if ($null -ne $val) { Escape-XmlText ($val.ToString()) } else { '' }
                    $colLetter = ConvertTo-ExcelColName $c
                    $sw.Write("<c r=`"$colLetter$rowNum`" s=`"0`" t=`"inlineStr`"><is><t>$valStr</t></is></c>")
                }
                $sw.Write("</row>")
            }

            $sw.Write("</sheetData></worksheet>")
        } finally { $sw.Dispose() }

        # Second sheet if requested
        if ($hasSecond) {
            $s2Entry = $zip.CreateEntry('xl/worksheets/sheet2.xml')
            $sw = New-Object System.IO.StreamWriter($s2Entry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><dimension ref=`"A1:C5`"/><sheetData><row r=`"1`"><c r=`"A1`" s=`"1`" t=`"inlineStr`"><is><t>Notatki / Notes</t></is></c><c r=`"B1`" s=`"1`" t=`"inlineStr`"><is><t>Data</t></is></c><c r=`"C1`" s=`"1`" t=`"inlineStr`"><is><t>Autor</t></is></c></row><row r=`"2`"><c r=`"A2`" s=`"2`" t=`"inlineStr`"><is><t>Ten arkusz musi przetrwac bez zmian!</t></is></c><c r=`"B2`" t=`"inlineStr`"><is><t>2026-10-01</t></is></c><c r=`"C2`" t=`"inlineStr`"><is><t>Admin</t></is></c></row></sheetData></worksheet>")
            } finally { $sw.Dispose() }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $fs) { $fs.Dispose() }
    }

    Move-Item -Path $tempZip -Destination $FilePath -Force
}

# --- Build Fixture Data ---
# Polish names & children fixture data (50 base records)
$firstNames = @('Jan','Anna','Piotr','Maria','Krzysztof','Katarzyna','Tomasz','Magdalena','Michal','Agnieszka','Pawel','Monika','Marcin','Ewa','Jakub','Zofia','Stanislaw','Aleksandra','Wojciech','Natalia','Adam','Karolina','Lukasz','Joanna','Mateusz','Paulina','Grzegorz','Dominika','Bartosz','Kinga','Kamil','Weronika','Szymon','Patrycja','Filip','Wiktoria','Daniel','Maja','Maciej','Oliwia','Kacper','Alicja','Artur','Julia','Robert','Emilia','Rafal','Zuzanna','Hubert','Iga')
$lastNames = @('Nowak','Kowalski','Wisniewski','Wojcik','Kowalczyk','Kaminski','Lewandowski','Zielinski','Szymanski','Wozniak','Dabrowski','Kozlowski','Jankowski','Mazur','Wojciechowski','Kwiatkowski','Krawczyk','Kaczmarek','Piotrowski','Grabowski','Zajac','Pawlak','Michalski','Krol','Wieczorek','Jablonski','Wrobel','Nowakowski','Majewski','Olszewski','Stepien','Malinowski','Jaworski','Adamczyk','Dudek','Nowicki','Pawlowski','Gorski','Witkowski','Walczak','Sikora','Baran','Rutkowski','Michalak','Szewczyk','Ostrowski','Tomaszewski','Pietrzak','Zalewski','Wroblewski')
$streets = @('Polna','Lesna','Sloneczna','Krotka','Szkolna','Ogrodowa','Brzozowa','Lipowa','Kwiatowa','Sosnowa')
$schools = @('SP nr 1','SP nr 3','SP nr 5','SP nr 12','Zespol Szkol Sportowych')
$vehicles = @('Bus 1 (Trasa A)','Bus 2 (Trasa B)','Bus 3 (Trasa C)','Bus 4 (Trasa D)','Wlasny transport')

$baseHeaders = @(
    'IdDziecka',
    'ImieNazwiskoDziecka',
    'DataUrodzenia',
    'PESEL',
    'Klasa',
    'Szkola',
    'ImieRodzica',
    'TelefonRodzica',
    'EmailRodzica',
    'Miejscowosc',
    'Ulica',
    'NumerDomu',
    'Pojazd',
    'Przystanek',
    'StatusOpieki',
    'Uwagi',
    'KategoriaDiety',
    'BiletMiesieczny',
    'ZgodaPrzetwarzanie',
    'OpiekunPrawny',
    'OstZmiana',
    'ZrodloSciezka',
    'ZrodloPlik',
    'Zmienil',
    'ZrodloWiersz',
    'ImportId'
)

$baseRows = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt 50; $i++) {
    $fn = $firstNames[$i]
    $ln = $lastNames[$i]
    $fullName = "$ln $fn"
    $pesel = "152$($i.ToString('D2'))12345"
    $street = $streets[$i % $streets.Count]
    $houseNo = ($i + 1).ToString()
    $school = $schools[$i % $schools.Count]
    $veh = $vehicles[$i % $vehicles.Count]

    $row = [ordered]@{
        IdDziecka           = "D-$($i + 1001)"
        ImieNazwiskoDziecka = $fullName
        DataUrodzenia       = "2015-0$((($i % 9) + 1))-15"
        PESEL               = $pesel
        Klasa               = "$((($i % 8) + 1))A"
        Szkola              = $school
        ImieRodzica         = "Rodzic $fullName"
        TelefonRodzica      = "500-111-$($i.ToString('D3'))"
        EmailRodzica        = "rodzic$($i + 1)@szkola.pl"
        Miejscowosc         = 'Warszawa'
        Ulica               = $street
        NumerDomu           = $houseNo
        Pojazd              = $veh
        Przystanek          = "$street / Glowna"
        StatusOpieki        = 'Aktywny'
        Uwagi               = 'Brak uwag'
        KategoriaDiety      = 'Standardowa'
        BiletMiesieczny     = 'Tak'
        ZgodaPrzetwarzanie  = 'Tak'
        OpiekunPrawny       = 'Oboje rodzicow'
        OstZmiana           = '2026-09-01 10:00'
        ZrodloSciezka       = 'C:\Dane\BazaPoczatkowa.xlsx'
        ZrodloPlik          = 'BazaPoczatkowa.xlsx'
        Zmienil             = 'System'
        ZrodloWiersz        = ($i + 2).ToString()
        ImportId            = '00000000-0000-0000-0000-000000000000'
    }
    [void]$baseRows.Add($row)
}

# --- Fixture 1: base.xlsx ---
$basePath = Join-Path $OutputDir 'base.xlsx'
New-SimpleExcelFile -FilePath $basePath -SheetName 'Dzieci' -Headers $baseHeaders -Rows $baseRows
Write-Host "Created base.xlsx at: $basePath"

# --- Fixture 2: base_formatted.xlsx ---
$baseFormattedPath = Join-Path $OutputDir 'base_formatted.xlsx'
New-SimpleExcelFile -FilePath $baseFormattedPath -SheetName 'Dzieci' -Headers $baseHeaders -Rows $baseRows -SecondSheetName 'Instrukcja'
Write-Host "Created base_formatted.xlsx at: $baseFormattedPath"

# --- Fixture 3: incoming_A.xlsx ---
# Different column order, address in ONE column ('PelnyAdres' = 'Miejscowosc, Ulica NumerDomu')
# 3 changed rows (row 5: phone change, row 10: school + vehicle change, row 20: diet change)
# 2 new rows (records 51 and 52)
# 45 unchanged rows (indices 0..49 except 4, 9, 19, plus remove 2 so total base matching = 48)
$incomingAHeaders = @(
    'KodUcznia',
    'NazwiskoIImie',
    'SzkolaPodstawowa',
    'OddzialKlasowy',
    'TelefonKontaktowy',
    'AdresZamieszkania',
    'SrodekTransportu',
    'UwagiDodatkowe',
    'WymogiDietetyczne'
)

$incomingARows = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt 50; $i++) {
    $fn = $firstNames[$i]
    $ln = $lastNames[$i]
    $fullName = "$ln $fn"
    $street = $streets[$i % $streets.Count]
    $houseNo = ($i + 1).ToString()
    $school = $schools[$i % $schools.Count]
    $veh = $vehicles[$i % $vehicles.Count]
    $phone = "500-111-$($i.ToString('D3'))"
    $diet = 'Standardowa'
    $notes = 'Brak uwag'

    # Introduce changes:
    if ($i -eq 4) {
        # Row 5 (index 4): phone number changed
        $phone = '600-999-888'
    } elseif ($i -eq 9) {
        # Row 10 (index 9): school and vehicle changed
        $school = 'SP nr 12'
        $veh = 'Bus 4 (Trasa D)'
    } elseif ($i -eq 19) {
        # Row 20 (index 19): diet changed
        $diet = 'Bezglutenowa'
    }

    # Skip index 48 and 49 so incoming has 48 existing (45 unchanged, 3 changed)
    if ($i -ge 48) { continue }

    $fullAddr = "Warszawa, $street $houseNo"

    $rowA = [ordered]@{
        KodUcznia          = "D-$($i + 1001)"
        NazwiskoIImie      = $fullName
        SzkolaPodstawowa   = $school
        OddzialKlasowy     = "$((($i % 8) + 1))A"
        TelefonKontaktowy  = $phone
        AdresZamieszkania  = $fullAddr
        SrodekTransportu   = $veh
        UwagiDodatkowe     = $notes
        WymogiDietetyczne  = $diet
    }
    [void]$incomingARows.Add($rowA)
}

# Add 2 NEW rows
$newRow1 = [ordered]@{
    KodUcznia          = 'D-1051'
    NazwiskoIImie      = 'Kalinowski Robert'
    SzkolaPodstawowa   = 'SP nr 3'
    OddzialKlasowy     = '2B'
    TelefonKontaktowy  = '501-222-333'
    AdresZamieszkania  = 'Warszawa, Polna 99'
    SrodekTransportu   = 'Bus 1 (Trasa A)'
    UwagiDodatkowe     = 'Nowy uczen od pazdziernika'
    WymogiDietetyczne  = 'Wegetarianska'
}
[void]$incomingARows.Add($newRow1)

$newRow2 = [ordered]@{
    KodUcznia          = 'D-1052'
    NazwiskoIImie      = 'Wisniewska Helena'
    SzkolaPodstawowa   = 'SP nr 5'
    OddzialKlasowy     = '1A'
    TelefonKontaktowy  = '502-333-444'
    AdresZamieszkania  = 'Warszawa, Lesna 12'
    SrodekTransportu   = 'Bus 2 (Trasa B)'
    UwagiDodatkowe     = 'Nowy zapis'
    WymogiDietetyczne  = 'Standardowa'
}
[void]$incomingARows.Add($newRow2)

$incomingAPath = Join-Path $OutputDir 'incoming_A.xlsx'
New-SimpleExcelFile -FilePath $incomingAPath -SheetName 'AktualizacjaWrzesien' -Headers $incomingAHeaders -Rows $incomingARows
Write-Host "Created incoming_A.xlsx at: $incomingAPath (45 Unchanged, 3 Changed, 2 New)"

# --- Fixture 4: incoming_B.csv ---
# Edge case CSV:
# 1 row with empty join key
# 1 duplicate incoming key
$incomingBCsvPath = Join-Path $OutputDir 'incoming_B.csv'
$csvContent = [System.Collections.ArrayList]::new()
[void]$csvContent.Add("KodDziecka;ImieNazwisko;Klasa;Telefon;Uwagi")
[void]$csvContent.Add("D-1001;Nowak Jan;1A;500-111-000;Pierwszy wiersz normalny")
[void]$csvContent.Add(";Kowalski Adam;2B;500-111-001;BLAD - puste ID")
[void]$csvContent.Add("D-1003;Wisniewski Piotr;3A;500-111-002;Poprawny")
[void]$csvContent.Add("D-1003;Wisniewski Piotr Duplikat;3A;500-999-999;BLAD - zduplikowany klucz w pliku")
[void]$csvContent.Add("D-1099;Zupelnie Nowy;1B;500-888-777;Nowy uczen")

[System.IO.File]::WriteAllLines($incomingBCsvPath, $csvContent, [System.Text.UTF8Encoding]::new($true))
Write-Host "Created incoming_B.csv at: $incomingBCsvPath"

# --- Fixture 5: incoming_C.xlsx ---
# Identical to incoming_A.xlsx but with the 3 changes and 2 new rows already present in base
# Will be generated or tested after acceptance in P6.
Write-Host "Test fixtures successfully created in $OutputDir."