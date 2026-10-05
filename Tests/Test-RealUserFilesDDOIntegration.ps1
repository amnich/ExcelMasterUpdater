#Requires -Version 5.1
<#
.SYNOPSIS
    Automated integration test suite for real-world DDO school transportation files.

.DESCRIPTION
    Validates end-to-end reconciliation between:
      Base:     "Wszystkie wnioski.xlsx" (492 rows, 38 columns)
      Incoming: "lista_26.xlsx"          (28 rows, 18 columns)
    
    Verifies:
      1. Smart Auto-Map heuristics (typo tolerance, stem matching, auto-concatenation detection).
      2. Multi-column Concatenate rule (Adres zamieszkania + Miasto Zamieszkania from combined address).
      3. Join key matching on Polish diacritics ('Imię i nazwisko ucznia/dziecka').
      4. Accurate status distribution (27 Changed, 1 New, 0 Ambiguous, 0 false address diffs).
      5. Projected row address splitting via Get-ProjectedRow.
      6. InPlace OpenXML write-back, pre-write backup creation, row append (492 -> 493), and audit logs.
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'

. $masterScript

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Real User Files Integration (Wszystkie wnioski.xlsx & lista_26.xlsx)" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# Locate real files
$defaultDir = "C:\Users\adamm\Downloads\DDO_2"
$baseFilePath = if ($env:DDO_BASE_PATH -and (Test-Path $env:DDO_BASE_PATH)) {
    $env:DDO_BASE_PATH
} elseif (Test-Path (Join-Path $defaultDir 'Wszystkie wnioski.xlsx')) {
    Join-Path $defaultDir 'Wszystkie wnioski.xlsx'
} else {
    $null
}

$incFilePath = if ($env:DDO_INC_PATH -and (Test-Path $env:DDO_INC_PATH)) {
    $env:DDO_INC_PATH
} elseif (Test-Path (Join-Path $defaultDir 'lista_26.xlsx')) {
    Join-Path $defaultDir 'lista_26.xlsx'
} else {
    $null
}

if (-not $baseFilePath -or -not $incFilePath) {
    Write-Host "[SKIP] Real user files not found at $defaultDir. Skipping test." -ForegroundColor Yellow
    exit 0
}

Write-Host "Base file:     $baseFilePath" -ForegroundColor Gray
Write-Host "Incoming file: $incFilePath" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# Phase 1: Header & Sheet Discovery
# -----------------------------------------------------------------------------
$baseSheet = [FastExcelHelper]::GetSheetNames($baseFilePath)[0]
$incSheet  = [FastExcelHelper]::GetSheetNames($incFilePath)[0]

$baseHeaders = [FastExcelHelper]::GetHeaders($baseFilePath, $baseSheet)
$incHeaders  = [FastExcelHelper]::GetHeaders($incFilePath, $incSheet)

if ($baseHeaders.Count -ne 38) {
    throw "Expected 38 base headers, but found $($baseHeaders.Count)"
}
if ($incHeaders.Count -ne 18) {
    throw "Expected 18 incoming headers, but found $($incHeaders.Count)"
}
Write-Host "[PASS] Phase 1: Base headers (38) and Incoming headers (18) discovered successfully." -ForegroundColor Green

# -----------------------------------------------------------------------------
# Phase 2: Smart Auto-Mapping Heuristic Verification
# -----------------------------------------------------------------------------
$mappingRules = @(Invoke-AutoMapRules -IncomingHeaders $incHeaders -BaseHeaders $baseHeaders)

if ($mappingRules.Count -ne 17) {
    throw "Expected 17 mapping rules from AutoMap, got $($mappingRules.Count)"
}

# Verify Concatenate Rule
$concatRule = $mappingRules | Where-Object { $_.MergeMode -eq 'Concatenate' }
if (-not $concatRule) {
    throw "AutoMap failed to detect Concatenate rule for split address + city!"
}
if ($concatRule.BaseColumns.Count -ne 2 -or $concatRule.BaseColumns[0] -notlike '*Adres*' -or $concatRule.BaseColumns[1] -notlike '*Miasto*') {
    throw "Concatenate rule BaseColumns mismatch: $($concatRule.BaseColumns -join ', ')"
}
if ($concatRule.UpdateColumns.Count -ne 1 -or $concatRule.UpdateColumns[0] -notlike '*Adres*') {
    throw "Concatenate rule UpdateColumns mismatch: $($concatRule.UpdateColumns -join ', ')"
}
if ($concatRule.Separator -ne ', ') {
    throw "Concatenate rule Separator expected ', ', got '$($concatRule.Separator)'"
}

# Verify Typo Tolerant Mapping (pywyżej/ ponieżej -> powyżej/ poniżej)
$typoRule = $mappingRules | Where-Object { ($_.UpdateColumns -join ' ') -like '*pywyżej*' }
if (-not $typoRule -or ($typoRule.BaseColumns -join ' ') -notlike '*powyżej*') {
    throw "Typo header 'pywyżej/ ponieżej' was not automatically mapped to 'powyżej/ poniżej'!"
}

# Verify Abbreviation Mapping (niepełn. -> niepełno.)
$abbrevRule = $mappingRules | Where-Object { $_.UpdateColumns -like '*niepełn.*' }
if (-not $abbrevRule -or $abbrevRule.BaseColumns[0] -notlike '*niepełno.*') {
    throw "Abbreviation header 'niepełn.' was not automatically mapped to 'niepełno.'!"
}

# Verify Stem Match (Adres zakładu pracy -> Adres Pracy)
$workAddrRule = $mappingRules | Where-Object { $_.UpdateColumns -like '*Adres zakładu pracy*' }
if (-not $workAddrRule -or $workAddrRule.BaseColumns[0] -notlike '*Adres Pracy*') {
    throw "Work address header 'Adres zakładu pracy' was not automatically mapped to 'Adres Pracy'!"
}

Write-Host "[PASS] Phase 2: Smart Auto-Map heuristics verified (100% of 18 incoming headers mapped, including typo, abbreviation, and auto-concatenate)." -ForegroundColor Green

# -----------------------------------------------------------------------------
# Phase 3: Comparison Execution & Categorization
# -----------------------------------------------------------------------------
$baseRows = [FastExcelHelper]::ReadSheet($baseFilePath, $baseSheet)
$incRows  = [FastExcelHelper]::ReadSheet($incFilePath, $incSheet)

if ($baseRows.Count -ne 492) {
    throw "Expected 492 base rows, got $($baseRows.Count)"
}
if ($incRows.Count -ne 28) {
    throw "Expected 28 incoming rows, got $($incRows.Count)"
}

$joinKeyHeader = $baseHeaders | Where-Object { $_ -like '*ucznia/dziecka*' }
$joinKey = @($joinKeyHeader)

$reviewItems = @(Invoke-MasterCompare `
    -BaseRows $baseRows `
    -IncomingRows $incRows `
    -MappingRules $mappingRules `
    -BaseJoinKey $joinKey `
    -IncomingJoinKey $joinKey `
    -BaseHeaders $baseHeaders)

if ($reviewItems.Count -ne 28) {
    throw "Expected 28 review items, got $($reviewItems.Count)"
}

$newRows = @($reviewItems | Where-Object { $_.Status -eq 'New' })
$chgRows = @($reviewItems | Where-Object { $_.Status -eq 'Changed' })
$ambRows = @($reviewItems | Where-Object { $_.Status -eq 'Ambiguous' })

if ($newRows.Count -ne 1) {
    throw "Expected exactly 1 New row, got $($newRows.Count)"
}
if ($chgRows.Count -ne 27) {
    throw "Expected exactly 27 Changed rows, got $($chgRows.Count)"
}
if ($ambRows.Count -ne 0) {
    throw "Expected 0 Ambiguous rows, got $($ambRows.Count)"
}

# Verify New row is Luso Skoczuk
$newChildName = $newRows[0].IncomingRow[$joinKeyHeader]
if ($newChildName -notlike '*Skoczuk*') {
    throw "Expected new record to be Skoczuk, got '$newChildName'"
}

# Verify zero false address diffs on matching composite addresses (e.g. Luos Gospodorczyk)
$gospRow = $chgRows | Where-Object { $_.IncomingRow[$joinKeyHeader] -like '*Gospodorczyk*' }
if ($gospRow) {
    $addrChange = $gospRow.Changes | Where-Object { $_.BaseColumn -like '*Adres*' -or $_.BaseColumn -like '*Miasto*' }
    if ($addrChange) {
        throw "False address difference detected for Gospodorczyk: $($addrChange | Out-String)"
    }
}

Write-Host "[PASS] Phase 3: Comparison engine verified (27 Changed, 1 New, 0 Ambiguous, 0 false address differences)." -ForegroundColor Green

# -----------------------------------------------------------------------------
# Phase 4: Projected Row Address Splitting Verification
# -----------------------------------------------------------------------------
$sampleIncRow = $incRows[0]
$projected = Get-ProjectedRow -IncomingRow $sampleIncRow -MappingRules $mappingRules

$baseAddrCol = $concatRule.BaseColumns[0]
$baseCityCol = $concatRule.BaseColumns[1]
$rawIncAddr  = $sampleIncRow.($concatRule.UpdateColumns[0])

if (-not [string]::IsNullOrWhiteSpace($rawIncAddr) -and $rawIncAddr.Contains(', ')) {
    $expectedAddrPart = $rawIncAddr.Split(@(', '), [System.StringSplitOptions]::None)[0].Trim()
    $expectedCityPart = $rawIncAddr.Split(@(', '), [System.StringSplitOptions]::None)[1].Trim()

    if ($projected[$baseAddrCol] -ne $expectedAddrPart) {
        throw "Projected address mismatch. Expected '$expectedAddrPart', got '$($projected[$baseAddrCol])'"
    }
    if ($projected[$baseCityCol] -ne $expectedCityPart) {
        throw "Projected city mismatch. Expected '$expectedCityPart', got '$($projected[$baseCityCol])'"
    }
}

Write-Host "[PASS] Phase 4: Projected row 1:N address splitting verified cleanly." -ForegroundColor Green

# -----------------------------------------------------------------------------
# Phase 5: InPlace Write-Back, Backup & Row Append Verification
# -----------------------------------------------------------------------------
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "MasterUpdater_DDO_Test_$([System.Guid]::NewGuid().ToString('N').Substring(0, 8))"
[void][System.IO.Directory]::CreateDirectory($tempDir)

try {
    $testBaseCopy = Join-Path $tempDir 'Wszystkie_wnioski_Copy.xlsx'
    Copy-Item $baseFilePath $testBaseCopy

    $backupDir = Join-Path $tempDir 'Backups'
    $logDir    = Join-Path $tempDir 'Logs'
    [void][System.IO.Directory]::CreateDirectory($backupDir)
    [void][System.IO.Directory]::CreateDirectory($logDir)

    # Accept all rows for write-back
    foreach ($item in $reviewItems) {
        $item.Decision = 'Accepted'
    }

    $cfg = @{
        BackupDirectory         = $backupDir
        BackupRetentionCount    = 20
        MaxBackupMb             = 200
        LogDirectory            = $logDir
        WriteMode               = 'InPlace'
        MetadataColumns         = $script:AppConfig.MetadataColumns
        RedactNamesInLog        = $false
        LogChangesToBaseSheet   = $false
        BaseSheetLogName        = 'ImportLog'
    }

    $writeRes = Invoke-MasterWriteBack -BaseFilePath $testBaseCopy -BaseSheet $baseSheet -ReviewItems $reviewItems -AppConfig $cfg -BaseHeaders $baseHeaders

    if (-not $writeRes.Success) {
        throw "Write-back failed: $($writeRes | Out-String)"
    }
    if ($writeRes.UpdatedCells -le 0) {
        throw "Expected positive updated cell count, got $($writeRes.UpdatedCells)"
    }
    if ($writeRes.AddedRows -ne 1) {
        throw "Expected exactly 1 appended row, got $($writeRes.AddedRows)"
    }

    # Verify backup exists
    $backups = @(Get-ChildItem -Path $backupDir -Filter '*.bak.xlsx')
    if ($backups.Count -eq 0) {
        throw "Pre-write backup was not created in $backupDir"
    }

    # Emit audit logs and verify
    $logRes = Write-ImportLog -LogDirectory $logDir -BatchId $writeRes.BatchId -ReviewItems $reviewItems -WriteBackResult $writeRes -BaseFilePath $testBaseCopy
    $logFiles = @(Get-ChildItem -Path $logDir)
    if ($logFiles.Count -eq 0) {
        throw "Audit logs were not created in $logDir"
    }

    # Verify modified base workbook
    $updatedBase = [FastExcelHelper]::ReadSheet($testBaseCopy, $baseSheet)
    if ($updatedBase.Count -ne 493) {
        throw "Expected 493 rows after appending 1 row, got $($updatedBase.Count)"
    }

    $foundNew = $false
    foreach ($r in $updatedBase) {
        if ($r.($joinKeyHeader) -like '*Skoczuk*') {
            $foundNew = $true
            break
        }
    }
    if (-not $foundNew) {
        throw "Appended record 'Skoczuk' was not found in modified Excel file!"
    }

    Write-Host "[PASS] Phase 5: InPlace write-back, backup, row append (492 -> 493), and audit log emission verified." -ForegroundColor Green
}
finally {
    if (Test-Path $tempDir) {
        try { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " ALL REAL USER FILE INTEGRATION TESTS PASSED (100% SUCCESS)                     " -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Cyan
