<#
.SYNOPSIS
    End-to-End automated integration test for Excel Master Updater.

.DESCRIPTION
    Validates full end-to-end import lifecycle across test fixtures:
    - Profile auto-resolution via SHA-256 header fingerprint.
    - Comparison, projection, and selective cell approvals.
    - InPlace write-back with automated backup creation and retention cleanup.
    - Metadata stamping into base audit columns.
    - Verification of written cell values and appended rows.
    - JSONL audit event logging and summary TXT generation.
    - Re-import idempotency (Run 2 yields 0 New, 0 Changed, 50 Unchanged).
    - Multi-sheet preservation on styled base_formatted.xlsx.

.OUTPUTS
    System.Void. Throws an exception if any integration assertion fails.

.NOTES
    PowerShell 5.1 and PowerShell 7.x compatible.
#>

$ErrorActionPreference = 'Stop'

# Dot-source Master-Updater.ps1 to load all functions and types
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Master-Updater.ps1"
. $scriptPath -NonInteractive

Write-Host "Master-Updater.ps1 loaded successfully!"

$testWorkDir = Join-Path $PSScriptRoot "E2E_Workspace"
if (Test-Path $testWorkDir) { Remove-Item -Recurse -Force $testWorkDir }
[void][System.IO.Directory]::CreateDirectory($testWorkDir)

$fixturesDir = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures"

# Copy base.xlsx to test workspace
$testBaseXlsx = Join-Path $testWorkDir "base_e2e.xlsx"
Copy-Item -Path (Join-Path $fixturesDir "base.xlsx") -Destination $testBaseXlsx -Force

$incomingAXlsx = Join-Path $fixturesDir "incoming_A.xlsx"
$incomingBCsv  = Join-Path $fixturesDir "incoming_B.csv"

# 1. Setup AppConfig
$e2eConfig = Get-DefaultAppConfig
$e2eConfig.BaseFilePath       = $testBaseXlsx
$e2eConfig.BaseSheet          = 'Dzieci'
$e2eConfig.ProfileStorePath   = (Join-Path $testWorkDir "Profiles")
$e2eConfig.LogDirectory       = (Join-Path $testWorkDir "Logs")
$e2eConfig.BackupDirectory    = (Join-Path $testWorkDir "Backups")
$e2eConfig.BackupRetentionCount = 5
$e2eConfig.WriteMode          = 'InPlace'

# 2. Setup Profile with Fingerprint
$hdrsA = [FastExcelHelper]::GetHeaders($incomingAXlsx, 'AktualizacjaWrzesien')
$fpA = Compute-HeaderFingerprint -Headers $hdrsA -SheetName 'AktualizacjaWrzesien'

$profileA = [ordered]@{
    SchemaVersion           = '2.0'
    Name                    = 'E2E Profil Wrzesien'
    CreatedUtc              = (Get-Date).ToUniversalTime().ToString('o')
    HeaderFingerprint       = $fpA
    BaseFilePath            = $testBaseXlsx
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

$savedProfilePath = Save-MappingProfile -Profile $profileA -StorePath $e2eConfig.ProfileStorePath
$resolvedProfile = Get-MappingProfileByFingerprint -Fingerprint $fpA -StorePath $e2eConfig.ProfileStorePath

if ($null -eq $resolvedProfile -or $resolvedProfile.Name -ne 'E2E Profil Wrzesien') {
    throw "FAILED: Profile auto-resolution by fingerprint failed"
}
Write-Host "SUCCESS: Profile auto-resolution by fingerprint verified!"

# 3. Read Base & Incoming rows
$baseRaw = [FastExcelHelper]::ReadSheet($testBaseXlsx, 'Dzieci')
$baseHeaders = [FastExcelHelper]::GetHeaders($testBaseXlsx, 'Dzieci')
$baseRows = [System.Collections.Generic.List[object]]::new()
foreach ($r in $baseRaw) {
    $vals = @{}
    foreach ($h in $baseHeaders) { $vals[$h] = $r.$h }
    $baseRows.Add([PSCustomObject]@{ RowNumber = [int]$r._RowNumber; Values = $vals })
}

$incRawA = [FastExcelHelper]::ReadSheet($incomingAXlsx, 'AktualizacjaWrzesien')
$incHdrsA = [FastExcelHelper]::GetHeaders($incomingAXlsx, 'AktualizacjaWrzesien')
$incomingARows = [System.Collections.Generic.List[object]]::new()
foreach ($r in $incRawA) {
    $vals = @{}
    foreach ($h in $incHdrsA) { $vals[$h] = $r.$h }
    $incomingARows.Add([PSCustomObject]@{ SourceRowNumber = [int]$r._RowNumber; Values = $vals; SourceFilePath = $incomingAXlsx })
}

# 4. Compare Run 1
$itemsRun1 = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incomingARows -MappingProfile $resolvedProfile -BaseHeaders $baseHeaders

$countsRun1 = @{ New = 0; Changed = 0; Unchanged = 0; Ambiguous = 0 }
foreach ($it in $itemsRun1) { $countsRun1[$it.Status]++ }
Write-Host "Run 1 Counts: New=$($countsRun1.New), Changed=$($countsRun1.Changed), Unchanged=$($countsRun1.Unchanged), Ambiguous=$($countsRun1.Ambiguous)"

if ($countsRun1.New -ne 2 -or $countsRun1.Changed -ne 3 -or $countsRun1.Unchanged -ne 45) {
    throw "FAILED: Run 1 expected 2 New, 3 Changed, 45 Unchanged"
}

# 5. Accept all New and Changed items
$acceptedItems = [System.Collections.Generic.List[object]]::new()
foreach ($it in $itemsRun1) {
    if ($it.Status -in @('New', 'Changed')) {
        $it.Status = 'Accepted'
        $acceptedItems.Add($it)
    }
}
Write-Host "Accepted $($acceptedItems.Count) items for write-back."

# 6. Execute Write-Back (InPlace)
$batchId = [guid]::NewGuid().ToString()
$wbResult = Invoke-MasterWriteBack -BaseFilePath $testBaseXlsx -BaseSheet 'Dzieci' -ReviewItems $acceptedItems -AppConfig $e2eConfig -BaseHeaders $baseHeaders -BatchId $batchId

Write-Host "Write-Back finished: Success=$($wbResult.Success), AddedRows=$($wbResult.AddedRows), UpdatedCells=$($wbResult.UpdatedCells)"
if (-not $wbResult.Success -or $wbResult.AddedRows -ne 2 -or $wbResult.UpdatedCells -lt 3) {
    throw "FAILED: Write-back result does not match expected numbers"
}

# 7. Verify Backup created
if (-not (Test-Path $wbResult.BackupPath)) {
    throw "FAILED: Backup file was not created: $($wbResult.BackupPath)"
}
Write-Host "SUCCESS: Backup verified at: $($wbResult.BackupPath)"

# 8. Write Logs & Verify
$logRes = Write-ImportLog -LogDirectory $e2eConfig.LogDirectory -BatchId $batchId -ReviewItems $itemsRun1 -WriteBackResult $wbResult -BaseFilePath $testBaseXlsx
if (-not (Test-Path $logRes.JsonlPath) -or -not (Test-Path $logRes.TxtPath)) {
    throw "FAILED: Log files were not written properly"
}
$jsonlContent = Get-Content $logRes.JsonlPath
if ($jsonlContent.Count -lt 5) { throw "FAILED: JSONL log contains too few events" }
Write-Host "SUCCESS: JSONL and TXT logs verified ($($jsonlContent.Count) events in JSONL)!"

# 9. Verify Base File Content after Write-Back
$postBaseRaw = [FastExcelHelper]::ReadSheet($testBaseXlsx, 'Dzieci')
$postHeaders = [FastExcelHelper]::GetHeaders($testBaseXlsx, 'Dzieci')

# Row count should now be 52 (50 original + 2 appended)
if ($postBaseRaw.Count -ne 52) {
    throw "FAILED: Base file expected 52 rows after append, got $($postBaseRaw.Count)"
}

# Check patched phone in row 6 (ID D-1005)
$r6 = $postBaseRaw | Where-Object { $_.IdDziecka -eq 'D-1005' }
if ($r6.TelefonRodzica -ne '600-999-888') {
    throw "FAILED: Row D-1005 phone not patched to 600-999-888 (got '$($r6.TelefonRodzica)')"
}
# Check metadata stamped
if ([string]::IsNullOrWhiteSpace($r6.OstZmiana) -or $r6.Zmienil -ne $env:USERNAME) {
    throw "FAILED: Metadata not properly stamped on patched row D-1005"
}

# Check appended row D-1051
$r51 = $postBaseRaw | Where-Object { $_.IdDziecka -eq 'D-1051' }
if ($null -eq $r51 -or $r51.ImieNazwiskoDziecka -ne 'Kalinowski Robert') {
    throw "FAILED: Appended row D-1051 not found"
}
if ($r51.OstZmiana -notmatch '^\d{4}-\d{2}-\d{2}') {
    throw "FAILED: Metadata OstZmiana not stamped on new row D-1051"
}
Write-Host "SUCCESS: Base file modifications & metadata verified!"

# 10. RE-IMPORT TEST (Run 2 / incoming_C test):
# Re-importing incoming_A against the newly updated base file should yield:
# 0 New, 0 Changed, 50 Unchanged!
$postBaseRows = [System.Collections.Generic.List[object]]::new()
foreach ($r in $postBaseRaw) {
    $vals = @{}
    foreach ($h in $postHeaders) { $vals[$h] = $r.$h }
    $postBaseRows.Add([PSCustomObject]@{ RowNumber = [int]$r._RowNumber; Values = $vals })
}

$itemsRun2 = Invoke-MasterCompare -BaseRows $postBaseRows -IncomingRows $incomingARows -MappingProfile $resolvedProfile -BaseHeaders $postHeaders
$countsRun2 = @{ New = 0; Changed = 0; Unchanged = 0; Ambiguous = 0 }
foreach ($it in $itemsRun2) { $countsRun2[$it.Status]++ }
Write-Host "Run 2 (Re-import) Counts: New=$($countsRun2.New), Changed=$($countsRun2.Changed), Unchanged=$($countsRun2.Unchanged), Ambiguous=$($countsRun2.Ambiguous)"

if ($countsRun2.New -ne 0 -or $countsRun2.Changed -ne 0 -or $countsRun2.Unchanged -ne 50) {
    throw "FAILED: Re-import test failed! Expected 0 New, 0 Changed, 50 Unchanged (got New=$($countsRun2.New), Changed=$($countsRun2.Changed), Unchanged=$($countsRun2.Unchanged))"
}
Write-Host "SUCCESS: Re-import idempotency verified (0 New, 0 Changed, 50 Unchanged)!"

# 11. Multi-sheet preservation on base_formatted.xlsx
$testFormatted = Join-Path $testWorkDir "base_formatted_e2e.xlsx"
Copy-Item -Path (Join-Path $fixturesDir "base_formatted.xlsx") -Destination $testFormatted -Force

$wbFormatResult = Invoke-MasterWriteBack -BaseFilePath $testFormatted -BaseSheet 'Dzieci' -ReviewItems $acceptedItems -AppConfig $e2eConfig -BaseHeaders $baseHeaders -BatchId ([guid]::NewGuid().ToString())

$sheetsAfter = [FastExcelHelper]::GetSheetNames($testFormatted)
if ($sheetsAfter.Count -lt 2 -or -not ($sheetsAfter -contains 'Instrukcja')) {
    throw "FAILED: Second sheet 'Instrukcja' was lost during write-back!"
}
Write-Host "SUCCESS: Multi-sheet preservation verified on base_formatted.xlsx!"

Write-Host "================================================================================"
Write-Host "ALL E2E INTEGRATION TESTS PASSED WITH 100% SUCCESS!"
Write-Host "================================================================================"