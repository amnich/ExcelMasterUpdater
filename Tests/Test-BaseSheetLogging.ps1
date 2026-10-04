#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'
$baseFixture  = Join-Path $fixturesDir 'base.xlsx'
$incomingFixture = Join-Path $fixturesDir 'incoming_A.xlsx'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Base Sheet Logging (ImportLog) Verification" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Dot-source Master-Updater.ps1
. $masterScript
Write-Host "[PASS] Master-Updater.ps1 dot-sourced successfully" -ForegroundColor Green

# 2. Test Get-DefaultAppConfig and config settings
$defCfg = Get-DefaultAppConfig
if (-not $defCfg.Contains('LogChangesToBaseSheet') -or $defCfg.LogChangesToBaseSheet -ne $false) {
    throw "DefaultAppConfig missing LogChangesToBaseSheet or not default `$false"
}
if (-not $defCfg.Contains('BaseSheetLogName') -or $defCfg.BaseSheetLogName -ne 'ImportLog') {
    throw "DefaultAppConfig missing BaseSheetLogName or not default 'ImportLog'"
}
Write-Host "[PASS] AppConfig defaults verified (LogChangesToBaseSheet=False, BaseSheetLogName='ImportLog')" -ForegroundColor Green

# Create isolated scratch workspace for tests
$scratchDir = Join-Path ([System.IO.Path]::GetTempPath()) ("BaseSheetLogTest_" + [System.Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($scratchDir)

try {
    # 3. Direct AppendWorksheetLog creation and appending test
    $testBase1 = Join-Path $scratchDir 'test_base1.xlsx'
    [System.IO.File]::Copy($baseFixture, $testBase1, $true)

    $testHeaders = @("Timestamp", "BatchId", "User", "Action", "BaseRow", "Key", "Column", "OldValue", "NewValue", "SourceFile", "SourceRow")
    $rowsBatch1 = [System.Collections.Generic.List[object]]::new()
    $rowsBatch1.Add(@('2026-10-04 08:00:00', 'BATCH-001', 'Admin', 'CellChanged', '2', 'K-101', 'Klasa', '3A', '4A', 'inc.xlsx', '5'))
    $rowsBatch1.Add(@('2026-10-04 08:00:00', 'BATCH-001', 'Admin', 'CellChanged', '2', 'K-101', 'Telefon', '111-222', '333-444', 'inc.xlsx', '5'))

    [EditExcelHelper]::AppendWorksheetLog($testBase1, 'ImportLog', $testHeaders, $rowsBatch1)

    # Verify sheet names via FastExcelHelper
    $sheets = [FastExcelHelper]::GetSheetNames($testBase1)
    if ($sheets -notcontains 'ImportLog') {
        throw "AppendWorksheetLog failed: 'ImportLog' sheet not found in workbook. Present: $($sheets -join ', ')"
    }
    Write-Host "[PASS] AppendWorksheetLog created new 'ImportLog' sheet in workbook" -ForegroundColor Green

    # Read sheet data and verify headers + rows
    $headers1 = [FastExcelHelper]::GetHeaders($testBase1, 'ImportLog')
    if ($headers1.Count -ne 11) {
        throw "Expected 11 columns in ImportLog headers, got $($headers1.Count)"
    }
    $logData1 = [FastExcelHelper]::ReadSheet($testBase1, 'ImportLog')
    if ($logData1.Count -ne 2) {
        throw "Expected 2 rows in ImportLog sheet, got $($logData1.Count)"
    }
    if ($logData1[0].Key -ne 'K-101' -or $logData1[0].Column -ne 'Klasa' -or $logData1[0].NewValue -ne '4A') {
        throw "First log row values do not match expected CellChanged data"
    }
    Write-Host "[PASS] ReadSheet verified 11 headers and 2 data rows in 'ImportLog'" -ForegroundColor Green

    # Append second batch to existing sheet
    $rowsBatch2 = [System.Collections.Generic.List[object]]::new()
    $rowsBatch2.Add(@('2026-10-04 08:05:00', 'BATCH-002', 'Admin', 'AddedRow', 'New', 'K-999', '(All)', '', 'Jan Kowalski; 1A', 'inc.xlsx', '12'))
    [EditExcelHelper]::AppendWorksheetLog($testBase1, 'ImportLog', $testHeaders, $rowsBatch2)

    $logData2 = [FastExcelHelper]::ReadSheet($testBase1, 'ImportLog')
    if ($logData2.Count -ne 3) {
        throw "Expected 3 total rows after second append batch, got $($logData2.Count)"
    }
    if ($logData2[2].Action -ne 'AddedRow' -or $logData2[2].Key -ne 'K-999') {
        throw "Third row values do not match second batch AddedRow data"
    }
    Write-Host "[PASS] Multi-batch append to existing sheet verified (total 3 rows, no duplicate headers)" -ForegroundColor Green

    # Verify original base data sheet integrity ('Dzieci')
    $dzieciData = [FastExcelHelper]::ReadSheet($testBase1, 'Dzieci')
    if ($dzieciData.Count -eq 0) {
        throw "Original 'Dzieci' data sheet was corrupted during ImportLog appending"
    }
    Write-Host "[PASS] Original data sheet 'Dzieci' integrity preserved intact" -ForegroundColor Green

    # 4. Custom Sheet Name & Sanitization Test
    $testBaseCustom = Join-Path $scratchDir 'test_custom.xlsx'
    [System.IO.File]::Copy($baseFixture, $testBaseCustom, $true)

    $customSheetRaw = "Audit:History/2026*New?"
    $cleanSheet = [regex]::Replace($customSheetRaw, '[\\/\?\*:[\]]', '_')
    if ($cleanSheet.Length -gt 31) { $cleanSheet = $cleanSheet.Substring(0, 31) }
    [EditExcelHelper]::AppendWorksheetLog($testBaseCustom, $cleanSheet, $testHeaders, $rowsBatch1)

    $customSheets = [FastExcelHelper]::GetSheetNames($testBaseCustom)
    if ($customSheets -notcontains $cleanSheet) {
        throw "Sanitized custom sheet '$cleanSheet' not found in workbook. Present: $($customSheets -join ', ')"
    }
    Write-Host "[PASS] Custom sheet name with invalid characters sanitized and created ('$cleanSheet')" -ForegroundColor Green

    # 5. Write-ImportLog integration with PII Masking and Sheet Logging
    $testBaseLogging = Join-Path $scratchDir 'test_logging.xlsx'
    [System.IO.File]::Copy($baseFixture, $testBaseLogging, $true)

    $fakeReviewItems = [System.Collections.Generic.List[object]]::new()
    $fakeReviewItem = [PSCustomObject]@{
        Status          = 'Accepted'
        Decision        = 'Accept'
        Title           = 'Kowalski Adam'
        MatchedBaseRow  = [PSCustomObject]@{
            RowNumber = 5
            Values    = [ordered]@{ 'Id' = 'ID-005'; 'Nazwisko' = 'Kowalski'; 'Klasa' = '2B' }
        }
        IncomingRow     = [PSCustomObject]@{
            SourceFilePath   = 'C:\Data\incoming.xlsx'
            SourceRowNumber  = 10
            Raw              = [ordered]@{ 'Id' = 'ID-005'; 'Nazwisko' = 'Kowalski'; 'Klasa' = '3B' }
        }
        Changes         = @(
            [PSCustomObject]@{
                BaseColumn        = 'Nazwisko'
                OldValue          = 'Kowalski'
                NewValue          = 'Nowak'
                CellRef           = 'B5'
                SelectedForUpdate = $true
            },
            [PSCustomObject]@{
                BaseColumn        = 'Klasa'
                OldValue          = '2B'
                NewValue          = '3B'
                CellRef           = 'C5'
                SelectedForUpdate = $true
            }
        )
    }
    $fakeReviewItems.Add($fakeReviewItem)

    $fakeWriteBackRes = @{
        BatchId          = 'BATCH-TEST-PII'
        UpdatedCells     = 2
        AddedRows        = 0
        PhysicallyDeleted = 0
        Timestamp        = (Get-Date)
    }

    $logDir = Join-Path $scratchDir 'Logs'
    [void][System.IO.Directory]::CreateDirectory($logDir)

    # 5A. Sheet logging disabled: verify sheet is NOT created
    $resNoLog = Write-ImportLog -LogDirectory $logDir -BatchId 'BATCH-NO-SHEET' -ReviewItems $fakeReviewItems `
        -WriteBackResult $fakeWriteBackRes -BaseFilePath $testBaseLogging -RedactNames $false `
        -LogChangesToBaseSheet $false -BaseSheetLogName 'ImportLog'

    if ($resNoLog.BaseSheetLog -ne $null) {
        throw "BaseSheetLog should be `$null when LogChangesToBaseSheet is `$false"
    }
    $sheetsBefore = [FastExcelHelper]::GetSheetNames($testBaseLogging)
    if ($sheetsBefore -contains 'ImportLog') {
        throw "'ImportLog' sheet should NOT exist when LogChangesToBaseSheet is `$false"
    }
    Write-Host "[PASS] Write-ImportLog does not create sheet when LogChangesToBaseSheet=False" -ForegroundColor Green

    # 5B. Sheet logging enabled WITH PII masking: verify sheet creation and redaction
    $resWithLog = Write-ImportLog -LogDirectory $logDir -BatchId 'BATCH-WITH-SHEET' -ReviewItems $fakeReviewItems `
        -WriteBackResult $fakeWriteBackRes -BaseFilePath $testBaseLogging -RedactNames $true `
        -LogChangesToBaseSheet $true -BaseSheetLogName 'ImportLog'

    if ($resWithLog.BaseSheetLog -ne 'ImportLog') {
        throw "Expected BaseSheetLog 'ImportLog', got '$($resWithLog.BaseSheetLog)'"
    }
    $sheetsAfter = [FastExcelHelper]::GetSheetNames($testBaseLogging)
    if ($sheetsAfter -notcontains 'ImportLog') {
        throw "'ImportLog' sheet was not created when LogChangesToBaseSheet is `$true"
    }

    $loggedRows = [FastExcelHelper]::ReadSheet($testBaseLogging, 'ImportLog')
    if ($loggedRows.Count -ne 2) {
        throw "Expected 2 rows in ImportLog, got $($loggedRows.Count)"
    }

    # Verify PII masking on 'Nazwisko' column
    $nazwiskoRow = $loggedRows | Where-Object { $_.Column -eq 'Nazwisko' }
    if (-not $nazwiskoRow) { throw "Log row for 'Nazwisko' not found" }
    if ($nazwiskoRow.OldValue -ne '***' -or $nazwiskoRow.NewValue -ne '***') {
        throw "PII column 'Nazwisko' was not masked with '***' in sheet log. Old='$($nazwiskoRow.OldValue)', New='$($nazwiskoRow.NewValue)'"
    }
    Write-Host "[PASS] PII masking ('***') verified on sensitive columns in base sheet log" -ForegroundColor Green

    # Verify non-PII column 'Klasa' was NOT masked
    $klasaRow = $loggedRows | Where-Object { $_.Column -eq 'Klasa' }
    if (-not $klasaRow) { throw "Log row for 'Klasa' not found" }
    if ($klasaRow.OldValue -ne '2B' -or $klasaRow.NewValue -ne '3B') {
        throw "Non-PII column 'Klasa' should preserve actual values. Got Old='$($klasaRow.OldValue)', New='$($klasaRow.NewValue)'"
    }
    Write-Host "[PASS] Non-PII column 'Klasa' preserved unmasked in base sheet log" -ForegroundColor Green

    # 6. CSV Base File Warning Test
    $testCsv = Join-Path $scratchDir 'test_base.csv'
    "Id,Nazwa,Wartosc`n1,Test,100" | Set-Content -Path $testCsv -Encoding utf8
    $warnMsg = $null
    $resCsv = Write-ImportLog -LogDirectory $logDir -BatchId 'BATCH-CSV' -ReviewItems $fakeReviewItems `
        -WriteBackResult $fakeWriteBackRes -BaseFilePath $testCsv -RedactNames $false `
        -LogChangesToBaseSheet $true -BaseSheetLogName 'ImportLog' -WarningVariable warnMsg

    if ($resCsv.BaseSheetLog -ne $null) {
        throw "BaseSheetLog should be `$null for CSV files"
    }
    Write-Host "[PASS] CSV base file safely skips sheet logging without error" -ForegroundColor Green

    # 7. End-to-End Headless Run with Base Sheet Logging
    $testBaseHeadless = Join-Path $scratchDir 'test_headless.xlsx'
    [System.IO.File]::Copy($baseFixture, $testBaseHeadless, $true)

    $customCfg = Get-DefaultAppConfig
    $customCfg.LogChangesToBaseSheet = $true
    $customCfg.BaseSheetLogName = 'AuditTrail'
    $customCfg.LogDirectory = $logDir
    $customCfg.BackupDirectory = Join-Path $scratchDir 'Backups'

    $headlessRes = Invoke-HeadlessMasterUpdater -BaseFilePath $testBaseHeadless -IncomingPath $incomingFixture `
        -BaseSheet 'Dzieci' -IncomingSheet 'AktualizacjaWrzesien' -Config $customCfg -AutoAccept AllNonAmbiguous -Quiet

    $headlessSheets = [FastExcelHelper]::GetSheetNames($testBaseHeadless)
    if ($headlessSheets -notcontains 'AuditTrail') {
        throw "Headless run failed to create 'AuditTrail' sheet in base file. Present: $($headlessSheets -join ', ')"
    }
    $auditData = [FastExcelHelper]::ReadSheet($testBaseHeadless, 'AuditTrail')
    if ($auditData.Count -eq 0) {
        throw "AuditTrail sheet was created but contains 0 log rows after headless run"
    }
    Write-Host "[PASS] Full Headless Master Update logged $($auditData.Count) rows to sheet 'AuditTrail'" -ForegroundColor Green

    Write-Host ""
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host " ALL BASE SHEET LOGGING TESTS PASSED (100% SUCCESS)      " -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Green

} finally {
    if (Test-Path $scratchDir) {
        try { [System.IO.Directory]::Delete($scratchDir, $true) } catch { }
    }
}
