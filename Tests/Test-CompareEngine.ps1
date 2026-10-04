# Test-CompareEngine.ps1
# Unit test for P1 (Projected Row) and P2 (Invoke-MasterCompare)

$ErrorActionPreference = 'Stop'

# Dot-source Master-Updater.ps1 to load production helpers and compare engine
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Master-Updater.ps1"
. $scriptPath -NonInteractive

Write-Host "Master-Updater.ps1 loaded successfully for Test-CompareEngine!"

# --- TEST EXECUTION AGAINST FIXTURES ---
$fixtureDir = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures"
$baseXlsx = Join-Path $fixtureDir "base.xlsx"
$incomingAXlsx = Join-Path $fixtureDir "incoming_A.xlsx"
$incomingBCsv = Join-Path $fixtureDir "incoming_B.csv"

# 1. Test Fingerprint
$hdrsA = [FastExcelHelper]::GetHeaders($incomingAXlsx, 'AktualizacjaWrzesien')
$fpA1 = Compute-HeaderFingerprint -Headers $hdrsA -SheetName 'AktualizacjaWrzesien'
$fpA2 = Compute-HeaderFingerprint -Headers $hdrsA -SheetName 'AktualizacjaWrzesien'
if ($fpA1 -ne $fpA2 -or [string]::IsNullOrEmpty($fpA1)) { throw "Fingerprint calculation failed" }
Write-Host "Fingerprint for incoming_A.xlsx: $fpA1 (Deterministic: OK)"

# 2. Test Get-ProjectedRow with 1:N split and N:1 merge
$sampleIncoming = @{
    KodUcznia         = 'D-1001'
    NazwiskoIImie     = 'Kowalski Jan'
    AdresZamieszkania = 'Warszawa, Polna 1'
    SzkolaPodstawowa  = 'SP nr 1'
}
$sampleRules = @(
    @{ BaseColumns = @('IdDziecka'); UpdateColumns = @('KodUcznia'); MergeMode = 'Exact' },
    @{ BaseColumns = @('ImieNazwiskoDziecka'); UpdateColumns = @('NazwiskoIImie'); MergeMode = 'Exact' },
    @{ BaseColumns = @('Miejscowosc', 'UlicaDom'); UpdateColumns = @('AdresZamieszkania'); SplitSeparator = ', ' },
    @{ BaseColumns = @('Szkola'); UpdateColumns = @('SzkolaPodstawowa'); MergeMode = 'Exact' }
)
$proj = Get-ProjectedRow -IncomingValues $sampleIncoming -MappingRules $sampleRules
if ($proj['IdDziecka'] -ne 'D-1001' -or $proj['Miejscowosc'] -ne 'Warszawa' -or $proj['UlicaDom'] -ne 'Polna 1') {
    throw "Get-ProjectedRow failed on sample test"
}
Write-Host "Get-ProjectedRow with 1:N delimiter split passed!"

# 3. Read base rows
$baseRaw = [FastExcelHelper]::ReadSheet($baseXlsx, 'Dzieci')
$baseHeaders = [FastExcelHelper]::GetHeaders($baseXlsx, 'Dzieci')
Write-Host "baseRaw count: $($baseRaw.Count), baseHeaders count: $($baseHeaders.Count)"
$baseRows = [System.Collections.Generic.List[object]]::new()
foreach ($r in $baseRaw) {
    $vals = @{}
    foreach ($h in $baseHeaders) {
        $vals[$h] = $r.$h
    }
    $baseRows.Add([PSCustomObject]@{
        RowNumber = [int]$r._RowNumber
        Values    = $vals
    })
}
Write-Host "baseRows count: $($baseRows.Count)"

# 4. Read incoming_A rows
$incRawA = [FastExcelHelper]::ReadSheet($incomingAXlsx, 'AktualizacjaWrzesien')
$incHdrsA = [FastExcelHelper]::GetHeaders($incomingAXlsx, 'AktualizacjaWrzesien')
$incomingARows = [System.Collections.Generic.List[object]]::new()
foreach ($r in $incRawA) {
    $vals = @{}
    foreach ($h in $incHdrsA) {
        $vals[$h] = $r.$h
    }
    $incomingARows.Add([PSCustomObject]@{
        SourceRowNumber = [int]$r._RowNumber
        Values          = $vals
        SourceFilePath  = $incomingAXlsx
    })
}

# Mapping Profile for incoming_A
$profileA = @{
    SchemaVersion           = '2.0'
    Name                    = 'Profil Wrzesien'
    HeaderFingerprint       = $fpA1
    BaseSheet               = 'Dzieci'
    SourceSheetName         = 'AktualizacjaWrzesien'
    JoinKeyBase             = @('IdDziecka')
    JoinKeyUpdate           = @('KodUcznia')
    EmptyIncomingMeansClear = $false
    CompareOptions          = @{
        IgnoreCase         = $true
        Trim               = $true
        IgnoreSpecialChars = $true
        IgnoreAllSpaces    = $true
    }
    MappingRules            = @(
        @{ BaseColumns = @('IdDziecka'); UpdateColumns = @('KodUcznia'); MergeMode = 'Exact' },
        @{ BaseColumns = @('ImieNazwiskoDziecka'); UpdateColumns = @('NazwiskoIImie'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Szkola'); UpdateColumns = @('SzkolaPodstawowa'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Klasa'); UpdateColumns = @('OddzialKlasowy'); MergeMode = 'Exact' },
        @{ BaseColumns = @('TelefonRodzica'); UpdateColumns = @('TelefonKontaktowy'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Pojazd'); UpdateColumns = @('SrodekTransportu'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Uwagi'); UpdateColumns = @('UwagiDodatkowe'); MergeMode = 'Exact' },
        @{ BaseColumns = @('KategoriaDiety'); UpdateColumns = @('WymogiDietetyczne'); MergeMode = 'Exact' }
    )
}

# Run Master Compare on incoming_A
$itemsA = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incomingARows -MappingProfile $profileA -BaseHeaders $baseHeaders

$countsA = @{ New = 0; Changed = 0; Unchanged = 0; Ambiguous = 0 }
foreach ($it in $itemsA) {
    $countsA[$it.Status]++
}

Write-Host "incoming_A comparison counts: New=$($countsA.New), Changed=$($countsA.Changed), Unchanged=$($countsA.Unchanged), Ambiguous=$($countsA.Ambiguous)"
$ambA = @($itemsA | Where-Object { $_.Status -eq 'Ambiguous' })
if ($ambA.Count -gt 0) {
    Write-Host "Sample Ambiguous reasons:"
    for ($i = 0; $i -lt [Math]::Min(5, $ambA.Count); $i++) {
        Write-Host "  #$($ambA[$i].Index): $($ambA[$i].AmbiguousReason) (MatchedBase: $($ambA[$i].MatchedBaseRow.RowNumber))"
    }
}

# Verify CellRef on changes
foreach ($it in $itemsA) {
    if ($it.Status -eq 'Changed') {
        foreach ($chg in $it.Changes) {
            Write-Host "  Change in BaseRow #$($it.MatchedBaseRow.RowNumber) Cell: $($chg.CellRef) Col: $($chg.BaseColumn) ('$($chg.OldValue)' -> '$($chg.NewValue)')"
            if ([string]::IsNullOrEmpty($chg.CellRef)) { throw "CellRef is empty!" }
        }
    }
}

# 5. Test incoming_B.csv edge cases
$csvLines = Get-Content $incomingBCsv
$csvHdrs = $csvLines[0].Split(';')
$incomingBRows = [System.Collections.Generic.List[object]]::new()
for ($l = 1; $l -lt $csvLines.Count; $l++) {
    $line = $csvLines[$l]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $cols = $line.Split(';')
    $vals = @{}
    for ($c = 0; $c -lt $csvHdrs.Length; $c++) {
        $vals[$csvHdrs[$c]] = if ($c -lt $cols.Length) { $cols[$c] } else { '' }
    }
    $incomingBRows.Add([PSCustomObject]@{
        SourceRowNumber = $l + 1
        Values          = $vals
        SourceFilePath  = $incomingBCsv
    })
}

$profileB = @{
    SchemaVersion           = '2.0'
    Name                    = 'Profil CSV B'
    JoinKeyBase             = @('IdDziecka')
    JoinKeyUpdate           = @('KodDziecka')
    EmptyIncomingMeansClear = $false
    CompareOptions          = @{
        IgnoreCase         = $true
        Trim               = $true
        IgnoreSpecialChars = $true
        IgnoreAllSpaces    = $true
    }
    MappingRules            = @(
        @{ BaseColumns = @('IdDziecka'); UpdateColumns = @('KodDziecka'); MergeMode = 'Exact' },
        @{ BaseColumns = @('ImieNazwiskoDziecka'); UpdateColumns = @('ImieNazwisko'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Klasa'); UpdateColumns = @('Klasa'); MergeMode = 'Exact' },
        @{ BaseColumns = @('TelefonRodzica'); UpdateColumns = @('Telefon'); MergeMode = 'Exact' },
        @{ BaseColumns = @('Uwagi'); UpdateColumns = @('Uwagi'); MergeMode = 'Exact' }
    )
}

$itemsB = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incomingBRows -MappingProfile $profileB -BaseHeaders $baseHeaders
$ambiguousCount = ($itemsB | Where-Object { $_.Status -eq 'Ambiguous' }).Count
Write-Host "incoming_B Ambiguous items: $ambiguousCount"
foreach ($it in ($itemsB | Where-Object { $_.Status -eq 'Ambiguous' })) {
    Write-Host "  Ambiguous Row #$($it.IncomingRow.SourceRowNumber): $($it.AmbiguousReason)"
}

if ($ambiguousCount -ne 2) {
    throw "FAILED: Expected exactly 2 ambiguous items in incoming_B.csv (empty key + duplicate key)"
}

# 6. Test DetectRemoved logic (Q3)
$itemsRemoved = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incomingARows -MappingProfile $profileA -BaseHeaders $baseHeaders -DetectRemoved $true -MarkDeletedColumn 'StatusOpieki' -MarkDeletedValue 'Usunięty'
$removedItems = @($itemsRemoved | Where-Object { $_.Status -eq 'Removed' })
Write-Host "incoming_A with DetectRemoved=true found $($removedItems.Count) removed items"
if ($removedItems.Count -ne 2) {
    throw "FAILED: Expected exactly 2 removed items when comparing incoming_A with DetectRemoved=true (got $($removedItems.Count))"
}
$remKeys = @($removedItems | ForEach-Object { $_.MatchedBaseRow.Values['IdDziecka'] })
Write-Host "Removed items detected: $($remKeys -join ', ')"
if (-not ($remKeys -contains 'D-1049' -and $remKeys -contains 'D-1050')) {
    throw "FAILED: Expected removed rows D-1049 and D-1050 (got $($remKeys -join ', '))"
}
if ($removedItems[0].Changes[0].NewValue -ne 'Usunięty') {
    throw "FAILED: Expected NewValue 'Usunięty' for removed item"
}
Write-Host "SUCCESS: DetectRemoved logic verified!"

Write-Host "ALL COMPARE ENGINE TESTS PASSED!"
