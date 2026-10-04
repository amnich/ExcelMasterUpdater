#Requires -Version 5.1
<#
.SYNOPSIS
    Generates realistic sample Excel (.xlsx) files for student school transport synchronization.

.DESCRIPTION
    Creates two matching OpenXML Excel workbooks:
    1. Baza_Uczniowie_Dowoz.xlsx - Master base roster containing 20 student transport records across 15 columns:
       Imię i Nazwisko Dziecka, Szkoła, Adres, Rodzaj dowozu, Opiekun, Telefon, Łączenie dowozu z pracą,
       Adres Praca, Samochód, Model, Pojemność, Uwagi, Ostatnia Zmiana, Plik zmiany, Aktualny.

    2. Zmiany_Wrzesien_2026.xlsx - Incoming change file containing 16 student transport records across 10 columns:
       Imię i Nazwisko Dziecka, Szkoła, Rodzaj dowozu, Opiekun, Telefon, Łączenie dowozu z pracą,
       Samochód, Model, Pojemność, Uwagi.

    Scenario distribution in Zmiany_Wrzesien_2026.xlsx:
    - 10 Unchanged records (exact matches with base roster)
    - 4 Changed records (phone, vehicle change, transport mode switch, school transfer)
    - 2 New records (new students enrolling in transport)

.PARAMETER OutputDir
    Target directory where generated sample files will be placed. Defaults to this script's directory.

.EXAMPLE
    pwsh -File .\Generate-SampleFiles.ps1
    Generates Baza_Uczniowie_Dowoz.xlsx and Zmiany_Wrzesien_2026.xlsx in .\SampleFiles.

.OUTPUTS
    System.IO.FileInfo. Details of created Excel files.

.NOTES
    Pure .NET System.IO.Compression and System.Xml. Zero Excel COM or third-party dependencies.
    Dual-compatible with Windows PowerShell 5.1 and PowerShell 7+.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

if (-not $OutputDir) {
    $OutputDir = $ScriptDir
}
if (-not (Test-Path $OutputDir)) {
    [void][System.IO.Directory]::CreateDirectory($OutputDir)
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml

function Escape-XmlText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}

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

function New-SampleExcelWorkbook {
    param(
        [string]$FilePath,
        [string]$SheetName = 'Arkusz1',
        [string[]]$Headers,
        [System.Collections.IList]$Rows,
        [string]$HeaderFillHex = 'FF1E3A8A' # Dark Blue default
    )

    if (Test-Path $FilePath) { Remove-Item -Force $FilePath }

    $tempZip = [System.IO.Path]::GetTempFileName()
    if (Test-Path $tempZip) { Remove-Item -Force $tempZip }

    $fs = $null
    $zip = $null
    try {
        $fs = [System.IO.File]::Create($tempZip)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

        # [Content_Types].xml
        $ctEntry = $zip.CreateEntry('[Content_Types].xml')
        $sw = New-Object System.IO.StreamWriter($ctEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/xl/workbook.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`"/><Override PartName=`"/xl/worksheets/sheet1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/><Override PartName=`"/xl/styles.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`"/></Types>")
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
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet1.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>")
        } finally { $sw.Dispose() }

        # xl/workbook.xml
        $wbEntry = $zip.CreateEntry('xl/workbook.xml')
        $sw = New-Object System.IO.StreamWriter($wbEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $sNameEsc = Escape-XmlText $SheetName
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets><sheet name=`"$sNameEsc`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>")
        } finally { $sw.Dispose() }

        # xl/styles.xml
        $stylesEntry = $zip.CreateEntry('xl/styles.xml')
        $sw = New-Object System.IO.StreamWriter($stylesEntry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><fonts count=`"2`"><font><sz val=`"11`"/><name val=`"Segoe UI`"/></font><font><b/><sz val=`"11`"/><color rgb=`"FFFFFFFF`"/><name val=`"Segoe UI`"/></font></fonts><fills count=`"3`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"$HeaderFillHex`"/></fill></fill></fills><borders count=`"2`"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=`"thin`"><color rgb=`"FFCBD5E1`"/></left><right style=`"thin`"><color rgb=`"FFCBD5E1`"/></right><top style=`"thin`"><color rgb=`"FFCBD5E1`"/></top><bottom style=`"thin`"><color rgb=`"FFCBD5E1`"/></bottom><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"2`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"1`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"2`" borderId=`"1`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`"/></cellXfs></styleSheet>")
        } finally { $sw.Dispose() }

        # xl/worksheets/sheet1.xml
        $s1Entry = $zip.CreateEntry('xl/worksheets/sheet1.xml')
        $sw = New-Object System.IO.StreamWriter($s1Entry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $totalRows = $Rows.Count + 1
            $lastColName = ConvertTo-ExcelColName ($Headers.Count - 1)
            $dimRef = "A1:$($lastColName)$($totalRows)"

            $sw.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><dimension ref=`"$dimRef`"/><sheetViews><sheetView tabSelected=`"1`" workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=`"18`"/><sheetData>")

            # Header row (s=1: bold white text on colored header)
            $sw.Write("<row r=`"1`" spans=`"1:$($Headers.Count)`" ht=`"22`" customHeight=`"1`">")
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $colLetter = ConvertTo-ExcelColName $c
                $hText = Escape-XmlText $Headers[$c]
                $sw.Write("<c r=`"$($colLetter)1`" s=`"1`" t=`"inlineStr`"><is><t>$hText</t></is></c>")
            }
            $sw.Write("</row>")

            # Data rows (s=0: bordered standard text)
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

        $zip.Dispose()
        $fs.Dispose()

        Copy-Item -Path $tempZip -Destination $FilePath -Force
    } finally {
        if ($zip) { $zip.Dispose() }
        if ($fs) { $fs.Dispose() }
        if (Test-Path $tempZip) { Remove-Item -Force $tempZip -ErrorAction SilentlyContinue }
    }
}

# -----------------------------------------------------------------------------
# 1. Base File Definition: Baza_Uczniowie_Dowoz.xlsx (15 columns)
# -----------------------------------------------------------------------------
$baseHeaders = @(
    'Imię i Nazwisko Dziecka',
    'Szkoła',
    'Adres',
    'Rodzaj dowozu',
    'Opiekun',
    'Telefon',
    'Łączenie dowozu z pracą',
    'Adres Praca',
    'Samochód',
    'Model',
    'Pojemność',
    'Uwagi',
    'Ostatnia Zmiana',
    'Plik zmiany',
    'Aktualny'
)

$baseData = @(
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kowalski Jan'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Kwiatowa 5, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kowalska Anna'
        'Telefon'                 = '601-234-567'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'Al. Jerozolimskie 100, Warszawa'
        'Samochód'                = 'Skoda'
        'Model'                   = 'Octavia'
        'Pojemność'               = '1968 cm³'
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Nowak Piotr'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Adres'                   = 'ul. Leśna 12, 05-820 Piastów'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Nowak Marek'
        'Telefon'                 = '602-345-678'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Porusza się na wózku inwalidzkim'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Wiśniewska Zuzanna'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Adres'                   = 'ul. Parkowa 3/4, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Wiśniewska Ewa'
        'Telefon'                 = '603-456-789'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Domaniewska 32, Warszawa'
        'Samochód'                = 'Toyota'
        'Model'                   = 'Yaris'
        'Pojemność'               = '1490 cm³'
        'Uwagi'                   = 'Wymaga stałej opieki'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Wójcik Michał'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Polna 8, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 1)'
        'Opiekun'                 = 'Wójcik Tomasz'
        'Telefon'                 = '604-567-890'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Przystanek: Polna/Główna'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kamińska Julia'
        'Szkoła'                  = 'Szkoła Podstawowa nr 3'
        'Adres'                   = 'ul. Słoneczna 15, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kamiński Robert'
        'Telefon'                 = '605-678-901'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Grójecka 50, Warszawa'
        'Samochód'                = 'Ford'
        'Model'                   = 'Focus'
        'Pojemność'               = '1498 cm³'
        'Uwagi'                   = 'Alergia wziewna'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Lewandowski Aleksander'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Adres'                   = 'ul. Brzozowa 22, 05-820 Piastów'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Lewandowska Monika'
        'Telefon'                 = '606-789-012'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = 'Opel'
        'Model'                   = 'Astra'
        'Pojemność'               = '1364 cm³'
        'Uwagi'                   = 'Wymaga asystenta ucznia'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Zielińska Maja'
        'Szkoła'                  = 'Szkoła Podstawowa nr 5'
        'Adres'                   = 'ul. Ogrodowa 7, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 2)'
        'Opiekun'                 = 'Zieliński Dariusz'
        'Telefon'                 = '607-890-123'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Przystanek: Ogrodowa PKP'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Szymański Bartosz'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Lipowa 18, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Szymańska Beata'
        'Telefon'                 = '608-901-234'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Kasprzaka 25, Warszawa'
        'Samochód'                = 'Volkswagen'
        'Model'                   = 'Golf'
        'Pojemność'               = '1598 cm³'
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Woźniak Filip'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Adres'                   = 'ul. Wierzbowa 4, 05-820 Piastów'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Woźniak Krzysztof'
        'Telefon'                 = '609-012-345'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Wózek spacerowy specjalny'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Dąbrowski Adam'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Adres'                   = 'ul. Długa 45, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 1)'
        'Opiekun'                 = 'Dąbrowska Katarzyna'
        'Telefon'                 = '610-123-456'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kozłowska Alicja'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Cicha 9, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kozłowski Wojciech'
        'Telefon'                 = '611-234-567'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Prosta 10, Warszawa'
        'Samochód'                = 'Hyundai'
        'Model'                   = 'i30'
        'Pojemność'               = '1396 cm³'
        'Uwagi'                   = 'Dieta bezmleczna'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Jankowski Mikołaj'
        'Szkoła'                  = 'Szkoła Podstawowa nr 3'
        'Adres'                   = 'ul. Spacerowa 11, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Jankowska Barbara'
        'Telefon'                 = '612-345-678'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Łopuszańska 36, Warszawa'
        'Samochód'                = 'Kia'
        'Model'                   = 'Ceed'
        'Pojemność'               = '1591 cm³'
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Mazur Wiktoria'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Adres'                   = 'ul. Kolejowa 16, 05-820 Piastów'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Mazur Andrzej'
        'Telefon'                 = '613-456-789'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Dziecko leżące / wózek'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kwiatkowski Antoni'
        'Szkoła'                  = 'Szkoła Podstawowa nr 5'
        'Adres'                   = 'ul. Stawowa 2, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 2)'
        'Opiekun'                 = 'Kwiatkowska Jolanta'
        'Telefon'                 = '614-567-890'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Przystanek: Kościół'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Krawczyk Laura'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Akacjowa 13, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Krawczyk Paweł'
        'Telefon'                 = '615-678-901'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Wołoska 12, Warszawa'
        'Samochód'                = 'Toyota'
        'Model'                   = 'Corolla'
        'Pojemność'               = '1798 cm³'
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Piotrowski Tymon'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Adres'                   = 'ul. Topolowa 21, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Piotrowska Danuta'
        'Telefon'                 = '616-789-012'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = 'Renault'
        'Model'                   = 'Megane'
        'Pojemność'               = '1461 cm³'
        'Uwagi'                   = 'Wymaga fotelika medycznego'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Grabowska Oliwia'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Adres'                   = 'ul. Wschodnia 30, 05-820 Piastów'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Grabowski Marcin'
        'Telefon'                 = '617-890-123'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Niedosłuch obustronny'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Nowakowski Szymon'
        'Szkoła'                  = 'Szkoła Podstawowa nr 3'
        'Adres'                   = 'ul. Południowa 6, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 1)'
        'Opiekun'                 = 'Nowakowska Magdalena'
        'Telefon'                 = '618-901-234'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Pawlak Natalia'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Adres'                   = 'ul. Zachodnia 17, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Pawlak Grzegorz'
        'Telefon'                 = '619-012-345'
        'Łączenie dowozu z pracą' = 'Tak'
        'Adres Praca'             = 'ul. Towarowa 28, Warszawa'
        'Samochód'                = 'Nissan'
        'Model'                   = 'Qashqai'
        'Pojemność'               = '1332 cm³'
        'Uwagi'                   = 'Brak uwag'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    },
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Michalska Maria'
        'Szkoła'                  = 'Szkoła Podstawowa nr 5'
        'Adres'                   = 'ul. Żytnia 3, 05-800 Pruszków'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 2)'
        'Opiekun'                 = 'Michalski Jarosław'
        'Telefon'                 = '620-123-456'
        'Łączenie dowozu z pracą' = 'Nie'
        'Adres Praca'             = ''
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Przystanek: Rondo'
        'Ostatnia Zmiana'         = '2026-08-20 10:15'
        'Plik zmiany'             = 'BazaPoczatkowa.xlsx'
        'Aktualny'                = 'Tak'
    }
)

# -----------------------------------------------------------------------------
# 2. Incoming Change File: Zmiany_Wrzesien_2026.xlsx (10 columns)
# -----------------------------------------------------------------------------
$incomingHeaders = @(
    'Imię i Nazwisko Dziecka',
    'Szkoła',
    'Rodzaj dowozu',
    'Opiekun',
    'Telefon',
    'Łączenie dowozu z pracą',
    'Samochód',
    'Model',
    'Pojemność',
    'Uwagi'
)

$incomingData = @(
    # Unchanged 1
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kowalski Jan'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kowalska Anna'
        'Telefon'                 = '601-234-567'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Skoda'
        'Model'                   = 'Octavia'
        'Pojemność'               = '1968 cm³'
        'Uwagi'                   = 'Brak uwag'
    },
    # Unchanged 2
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Nowak Piotr'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Nowak Marek'
        'Telefon'                 = '602-345-678'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Porusza się na wózku inwalidzkim'
    },
    # Unchanged 3
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Wiśniewska Zuzanna'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Wiśniewska Ewa'
        'Telefon'                 = '603-456-789'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Toyota'
        'Model'                   = 'Yaris'
        'Pojemność'               = '1490 cm³'
        'Uwagi'                   = 'Wymaga stałej opieki'
    },
    # Changed 1: Wójcik Michał switches from Gminny bus to Własny transport, adds car, new phone
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Wójcik Michał'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Wójcik Tomasz'
        'Telefon'                 = '501-999-111'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Toyota'
        'Model'                   = 'RAV4'
        'Pojemność'               = '2487 cm³'
        'Uwagi'                   = 'Zmiana na dowóz własny od 01.09.2026'
    },
    # Unchanged 4
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kamińska Julia'
        'Szkoła'                  = 'Szkoła Podstawowa nr 3'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kamiński Robert'
        'Telefon'                 = '605-678-901'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Ford'
        'Model'                   = 'Focus'
        'Pojemność'               = '1498 cm³'
        'Uwagi'                   = 'Alergia wziewna'
    },
    # Changed 2: Lewandowski Aleksander changed vehicle (Opel Astra -> Dacia Jogger)
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Lewandowski Aleksander'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Lewandowska Monika'
        'Telefon'                 = '606-789-012'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = 'Dacia'
        'Model'                   = 'Jogger'
        'Pojemność'               = '999 cm³'
        'Uwagi'                   = 'Wymaga asystenta ucznia (potwierdzone orzeczenie 2026/27)'
    },
    # Unchanged 5
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Zielińska Maja'
        'Szkoła'                  = 'Szkoła Podstawowa nr 5'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 2)'
        'Opiekun'                 = 'Zieliński Dariusz'
        'Telefon'                 = '607-890-123'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Przystanek: Ogrodowa PKP'
    },
    # Unchanged 6
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Szymański Bartosz'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Szymańska Beata'
        'Telefon'                 = '608-901-234'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Volkswagen'
        'Model'                   = 'Golf'
        'Pojemność'               = '1598 cm³'
        'Uwagi'                   = 'Brak uwag'
    },
    # Unchanged 7
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Woźniak Filip'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Woźniak Krzysztof'
        'Telefon'                 = '609-012-345'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Wózek spacerowy specjalny'
    },
    # Unchanged 8
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Dąbrowski Adam'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Rodzaj dowozu'           = 'Gminny bus szkolny (Trasa 1)'
        'Opiekun'                 = 'Dąbrowska Katarzyna'
        'Telefon'                 = '610-123-456'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Brak uwag'
    },
    # Changed 3: Kozłowska Alicja phone updated & work combining changed
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Kozłowska Alicja'
        'Szkoła'                  = 'SP nr 1 im. Mikołaja Kopernika'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Kozłowski Wojciech'
        'Telefon'                 = '533-777-888'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = 'Hyundai'
        'Model'                   = 'i30'
        'Pojemność'               = '1396 cm³'
        'Uwagi'                   = 'Dieta bezmleczna'
    },
    # Unchanged 9
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Jankowski Mikołaj'
        'Szkoła'                  = 'Szkoła Podstawowa nr 3'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Jankowska Barbara'
        'Telefon'                 = '612-345-678'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Kia'
        'Model'                   = 'Ceed'
        'Pojemność'               = '1591 cm³'
        'Uwagi'                   = 'Brak uwag'
    },
    # Unchanged 10
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Mazur Wiktoria'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Mazur Andrzej'
        'Telefon'                 = '613-456-789'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Dziecko leżące / wózek'
    },
    # Changed 4: Krawczyk Laura transferred to new school
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Krawczyk Laura'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Krawczyk Paweł'
        'Telefon'                 = '615-678-901'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Toyota'
        'Model'                   = 'Corolla'
        'Pojemność'               = '1798 cm³'
        'Uwagi'                   = 'Przeniesienie do ZSS nr 2 od września'
    },
    # New 1: Borkowski Stanisław
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Borkowski Stanisław'
        'Szkoła'                  = 'Szkoła Podstawowa nr 4'
        'Rodzaj dowozu'           = 'Własny transport rodzica'
        'Opiekun'                 = 'Borkowska Wioletta'
        'Telefon'                 = '721-333-444'
        'Łączenie dowozu z pracą' = 'Tak'
        'Samochód'                = 'Peugeot'
        'Model'                   = '3008'
        'Pojemność'               = '1199 cm³'
        'Uwagi'                   = 'Nowe zgłoszenie od 01.09.2026'
    },
    # New 2: Czarnecka Helena
    [ordered]@{
        'Imię i Nazwisko Dziecka' = 'Czarnecka Helena'
        'Szkoła'                  = 'Zespół Szkół Specjalnych nr 2'
        'Rodzaj dowozu'           = 'Przewoźnik specjalistyczny'
        'Opiekun'                 = 'Czarnecki Artur'
        'Telefon'                 = '722-555-666'
        'Łączenie dowozu z pracą' = 'Nie'
        'Samochód'                = ''
        'Model'                   = ''
        'Pojemność'               = ''
        'Uwagi'                   = 'Wózek aktywny, nowe orzeczenie'
    }
)

# Output Paths
$basePath = Join-Path $OutputDir 'Baza_Uczniowie_Dowoz.xlsx'
$incPath  = Join-Path $OutputDir 'Zmiany_Wrzesien_2026.xlsx'

Write-Host "Generating Base Excel file ($($baseData.Count) rows, $($baseHeaders.Count) cols)..." -ForegroundColor Cyan
New-SampleExcelWorkbook -FilePath $basePath -SheetName 'Dowóz Uczniów' -Headers $baseHeaders -Rows $baseData -HeaderFillHex 'FF1E3A8A'

Write-Host "Generating Incoming Changes Excel file ($($incomingData.Count) rows, $($incomingHeaders.Count) cols)..." -ForegroundColor Cyan
New-SampleExcelWorkbook -FilePath $incPath -SheetName 'Wnioski Wrzesień' -Headers $incomingHeaders -Rows $incomingData -HeaderFillHex 'FF065F46'

$bInfo = Get-Item $basePath
$iInfo = Get-Item $incPath

Write-Host "`nSUCCESS: Sample files generated successfully!" -ForegroundColor Green
Write-Host "  Base:     $($bInfo.FullName) ($([Math]::Round($bInfo.Length/1KB, 1)) KB)" -ForegroundColor White
Write-Host "  Incoming: $($iInfo.FullName) ($([Math]::Round($iInfo.Length/1KB, 1)) KB)" -ForegroundColor White

return @($bInfo, $iInfo)
