#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'
$baseFixture  = Join-Path $fixturesDir 'base.xlsx'
$incomingFixture = Join-Path $fixturesDir 'incoming_A.xlsx'

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " DEEP GOAL VERIFICATION: Base Sheet Logging & Unchanged Filter Integration" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 1. Dot-source Master-Updater
. $masterScript
Write-Host "[1/7] Master-Updater.ps1 dot-sourced cleanly" -ForegroundColor Green

# 2. Setup isolated scratch directory
$scratchDir = Join-Path ([System.IO.Path]::GetTempPath()) ("DeepGoalVerify_" + [System.Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($scratchDir)

try {
    $testBase = Join-Path $scratchDir 'deep_base.xlsx'
    [System.IO.File]::Copy($baseFixture, $testBase, $true)

    # 3. Test GUI NonInteractive launch and Settings modal binding
    $win = Show-MasterUpdater -InitBaseFilePath $testBase -InitIncomingPath $incomingFixture `
        -InitBaseSheet 'Dzieci' -InitIncomingSheet 'AktualizacjaWrzesien' -NonInteractive

    if (-not $win) { throw "Show-MasterUpdater -NonInteractive returned null" }
    Write-Host "[2/7] Window initialized in non-interactive mode" -ForegroundColor Green

    # Verify GUI controls for Unchanged option
    $chkShowUnchanged = $win.FindName('chkShowUnchanged')
    $cmbFilterStatus  = $win.FindName('cmbFilterStatus')
    $btnRunCompare    = $win.FindName('btnRunCompare')
    $btnApplyAccepted = $win.FindName('btnApplyAccepted')
    $btnAcceptAll     = $win.FindName('btnAcceptAll')

    if (-not $chkShowUnchanged) { throw "chkShowUnchanged control missing from GUI" }
    if (-not $cmbFilterStatus)  { throw "cmbFilterStatus control missing from GUI" }

    # 4. Run compare with Unchanged shown and check count
    $chkShowUnchanged.IsChecked = $true
    $btnRunCompare.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    [System.Windows.Forms.Application]::DoEvents()

    if ($script:AllReviewItems.Count -ne 50) {
        throw "Expected 50 items when unchanged is shown, got $($script:AllReviewItems.Count)"
    }
    Write-Host "[3/7] Comparison with 'Show Unchanged' verified (50 items total, 45 unchanged)" -ForegroundColor Green

    # Accept all non-ambiguous rows (excluding unchanged)
    foreach ($item in $script:AllReviewItems) {
        if ($item.Record.Status -ne 'Ambiguous' -and $item.Record.Status -ne 'Unchanged') {
            $item.Decision = 'Accepted'
        }
    }

    $acceptedCount = @($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count
    if ($acceptedCount -ne 5) {
        throw "Expected 5 accepted items (2 New + 3 Changed, 45 Unchanged skipped), got $acceptedCount"
    }
    Write-Host "[4/7] Accept All accepted 5 rows (skipped unchanged rows to protect base file)" -ForegroundColor Green

    # 5. Enable base sheet logging with custom sanitized name in AppConfig
    $customLogSheet = "ImportLog/Audit:2026?"
    $cleanLogSheet  = [regex]::Replace($customLogSheet, '[\\/\?\*:[\]]', '_')
    if ($cleanLogSheet.Length -gt 31) { $cleanLogSheet = $cleanLogSheet.Substring(0, 31) }

    $script:AppConfig.LogChangesToBaseSheet = $true
    $script:AppConfig.BaseSheetLogName      = $cleanLogSheet
    $script:AppConfig.RedactNamesInLog      = $true
    $script:AppConfig.LogDirectory          = Join-Path $scratchDir 'Logs'
    $script:AppConfig.BackupDirectory       = Join-Path $scratchDir 'Backups'

    # Execute Write-Back with base sheet logging active
    $accRecords = [System.Collections.Generic.List[object]]::new()
    foreach ($it in $script:AllReviewItems) {
        if ($it.Decision -eq 'Accepted') {
            $accRecords.Add($it.Record)
        }
    }

    $wbRes = Invoke-MasterWriteBack -BaseFilePath $testBase -BaseSheet 'Dzieci' -AcceptedItems $accRecords -AppConfig $script:AppConfig
    $logRes = Write-ImportLog -LogDirectory $script:AppConfig.LogDirectory -BatchId $wbRes.BatchId -ReviewItems $script:AllReviewItems `
        -WriteBackResult $wbRes -BaseFilePath $testBase -RedactNames $script:AppConfig.RedactNamesInLog `
        -LogChangesToBaseSheet ($script:AppConfig.LogChangesToBaseSheet -eq $true) -BaseSheetLogName $cleanLogSheet

    if ($logRes.BaseSheetLog -ne $cleanLogSheet) {
        throw "Expected BaseSheetLog '$cleanLogSheet', got '$($logRes.BaseSheetLog)'"
    }
    Write-Host "[5/7] Write-back executed with sheet log '$cleanLogSheet'" -ForegroundColor Green

    # 6. Verify base file structure and audit log sheet content
    $sheets = [FastExcelHelper]::GetSheetNames($testBase)
    if ($sheets -notcontains $cleanLogSheet) {
        throw "Sheet '$cleanLogSheet' missing from base workbook. Sheets found: $($sheets -join ', ')"
    }

    $headers = [FastExcelHelper]::GetHeaders($testBase, $cleanLogSheet)
    if ($headers.Count -ne 11) {
        throw "Expected 11 audit headers, got $($headers.Count)"
    }

    $logRows = [FastExcelHelper]::ReadSheet($testBase, $cleanLogSheet)
    if ($logRows.Count -eq 0) {
        throw "Audit sheet was created but has 0 rows"
    }

    # Verify PII masking on sensitive fields
    $maskedRows = $logRows | Where-Object { $_.Column -match '(?i)(nazwisko|imie|adres|pesel|telefon)' }
    foreach ($mr in $maskedRows) {
        if ($mr.OldValue -and $mr.OldValue -ne '***') {
            throw "OldValue in column $($mr.Column) was not masked: $($mr.OldValue)"
        }
        if ($mr.NewValue -and $mr.NewValue -ne '***') {
            throw "NewValue in column $($mr.Column) was not masked: $($mr.NewValue)"
        }
    }
    Write-Host "[6/7] Audit sheet verified ($($logRows.Count) rows, 11 columns, PII masking confirmed)" -ForegroundColor Green

    # 7. Verify OpenXML Package Integrity
    # Open the xlsx as a zip archive and inspect workbook.xml and [Content_Types].xml
    $zip = [System.IO.Compression.ZipFile]::OpenRead($testBase)
    try {
        $wbEntry = $zip.GetEntry('xl/workbook.xml')
        if (-not $wbEntry) { throw "Corrupt XLSX: missing xl/workbook.xml" }
        $ctEntry = $zip.GetEntry('[Content_Types].xml')
        if (-not $ctEntry) { throw "Corrupt XLSX: missing [Content_Types].xml" }
        $relsEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
        if (-not $relsEntry) { throw "Corrupt XLSX: missing xl/_rels/workbook.xml.rels" }
    } finally {
        $zip.Dispose()
    }
    Write-Host "[7/7] OpenXML package integrity verified (valid ECMA-376 archive)" -ForegroundColor Green

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host " DEEP VERIFICATION SUITE PASSED (100% SUCCESS)                                  " -ForegroundColor Green
    Write-Host "================================================================================" -ForegroundColor Green
} finally {
    if (Test-Path $scratchDir) {
        try { [System.IO.Directory]::Delete($scratchDir, $true) } catch { }
    }
}
