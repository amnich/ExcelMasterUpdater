#Requires -Version 5.1
<#
.SYNOPSIS
    Test suite for Headless Batch Pipeline, HTML Diff Reporting, and Folder Watcher (SWOT W3, O1, O2, O4).

.DESCRIPTION
    Validates:
    1. Standalone HTML diff report generation (Export-ReviewReport).
    2. Headless batch execution without UI (Invoke-HeadlessMasterUpdater -AutoAccept None).
    3. Headless automated write-back (Invoke-HeadlessMasterUpdater -AutoAccept AllNonAmbiguous).
    4. Folder-watch daemon processing (-Once mode), archiving to Archive/ folder.
    5. Folder-watch error routing to Failed/ folder on corrupt incoming file.

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir

$PassCount = 0
$FailCount = 0

function Test-Case {
    param([string]$Name, [scriptblock]$Block)
    try {
        & $Block
        Write-Host "[PASS] $Name" -ForegroundColor Green
        $script:PassCount++
    } catch {
        Write-Host "[FAIL] $Name -- $($_.Exception.Message)" -ForegroundColor Red
        $script:FailCount++
    }
}

function Assert {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Headless Mode, HTML Report & Folder Watcher" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# Prepare isolated test directory
$TestWorkspace = Join-Path $env:TEMP 'MasterUpdater_HeadlessTests'
if (Test-Path $TestWorkspace) { Remove-Item -Recurse -Force $TestWorkspace }
[void][System.IO.Directory]::CreateDirectory($TestWorkspace)

# Dot-source Master-Updater
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
. $masterScript -NonInteractive

# Copy test fixtures into workspace
$fixturesDir = Join-Path $ProjectRoot 'TestFixtures'
$baseOriginal = Join-Path $fixturesDir 'base.xlsx'
$incAOriginal = Join-Path $fixturesDir 'incoming_A.xlsx'
$incBOriginal = Join-Path $fixturesDir 'incoming_B.csv'

$baseTest = Join-Path $TestWorkspace 'base_test.xlsx'
$incTestA = Join-Path $TestWorkspace 'incoming_A.xlsx'
$incTestB = Join-Path $TestWorkspace 'incoming_B.csv'

Copy-Item $baseOriginal $baseTest
Copy-Item $incAOriginal $incTestA
Copy-Item $incBOriginal $incTestB

# Register mapping profile for incoming_A in ProfileStorePath
$cfg = Get-AppConfig
$incHdrsA = [FastExcelHelper]::GetHeaders($incAOriginal, 'AktualizacjaWrzesien')
$fpA1 = Compute-HeaderFingerprint -Headers $incHdrsA -SheetName 'AktualizacjaWrzesien'
$fpA2 = Compute-HeaderFingerprint -Headers $incHdrsA

$profileA = @{
    SchemaVersion           = '2.0'
    Name                    = 'Profil Wrzesien Test'
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
[void](Save-MappingProfile -Profile $profileA -StorePath $cfg.ProfileStorePath)
$profileA.HeaderFingerprint = $fpA2
[void](Save-MappingProfile -Profile $profileA -StorePath $cfg.ProfileStorePath)

# Test 1: HTML Diff Report Generation
Test-Case 'HTML Diff Report generates valid, standalone HTML with styled KPIs' {
    $htmlReportPath = Join-Path $TestWorkspace 'diff_report.html'
    $mockItems = [System.Collections.Generic.List[object]]::new()
    $mockItems.Add([PSCustomObject]@{
        Status   = 'Changed'
        Decision = 'Accepted'
        IndexStr = '#1'
        Title    = 'D-1004'
        Record   = [PSCustomObject]@{
            Status          = 'Changed'
            Changes         = @(
                [PSCustomObject]@{ BaseColumn = 'TelefonRodzica'; OldValue = '500-111-004'; NewValue = '600-999-888' }
            )
            MatchedBaseRow  = [PSCustomObject]@{ RowNumber = 6 }
            IncomingRow     = [PSCustomObject]@{ SourceRowNumber = 5 }
            AmbiguousReason = ''
        }
    })
    $mockItems.Add([PSCustomObject]@{
        Status   = 'New'
        Decision = 'Accepted'
        IndexStr = '#2'
        Title    = 'D-1051'
        Record   = [PSCustomObject]@{
            Status          = 'New'
            Changes         = @()
            MatchedBaseRow  = $null
            IncomingRow     = [PSCustomObject]@{ SourceRowNumber = 51 }
            AmbiguousReason = ''
        }
    })

    $resPath = Export-ReviewReport -ReviewItems $mockItems -DestinationPath $htmlReportPath -BaseFilePath $baseTest -IncomingPath $incTestA
    Assert (Test-Path $htmlReportPath) 'HTML report file must exist'

    $htmlContent = [System.IO.File]::ReadAllText($htmlReportPath)
    Assert ($htmlContent -like '*<!DOCTYPE html>*') 'HTML report must contain DOCTYPE'
    Assert ($htmlContent -like '*TelefonRodzica*') 'HTML report must contain column change name'
    Assert ($htmlContent -like '*600-999-888*') 'HTML report must contain new value'
    Assert ($htmlContent -like '*kpi-card*') 'HTML report must contain KPI styling'
}

# Test 2: Headless Comparison without Write-Back (-AutoAccept None)
Test-Case 'Headless comparison with AutoAccept None computes diff and leaves base file untouched' {
    $reportOut = Join-Path $TestWorkspace 'headless_report.html'
    $baseBefore = [System.IO.File]::GetLastWriteTimeUtc($baseTest)

    $res = Invoke-HeadlessMasterUpdater -BaseFilePath $baseTest -IncomingPath $incTestA -AutoAccept None -ExportReportPath $reportOut -Quiet

    Assert ($res.Success -eq $true) 'Invoke-HeadlessMasterUpdater must return Success=true'
    Assert ($res.TotalIncoming -eq 50) "Expected 50 incoming rows, got $($res.TotalIncoming)"
    Assert ($res.NewRows -eq 2) "Expected 2 new rows, got $($res.NewRows)"
    Assert ($res.ChangedRows -eq 3) "Expected 3 changed rows, got $($res.ChangedRows)"
    Assert ($res.AcceptedRows -eq 0) "AutoAccept None must accept 0 rows, got $($res.AcceptedRows)"
    Assert ($res.UpdatedCells -eq 0) 'No cells should be updated'
    Assert ($res.AddedRows -eq 0) 'No rows should be added'
    Assert (Test-Path $reportOut) 'Report must be exported'

    $baseAfter = [System.IO.File]::GetLastWriteTimeUtc($baseTest)
    Assert ($baseBefore -eq $baseAfter) 'Base file timestamp must remain unchanged with AutoAccept None'
}

# Test 3: Headless Automated Write-Back (-AutoAccept AllNonAmbiguous)
Test-Case 'Headless AutoAccept AllNonAmbiguous executes atomic write-back and creates backup/log' {
    $res = Invoke-HeadlessMasterUpdater -BaseFilePath $baseTest -IncomingPath $incTestA -AutoAccept AllNonAmbiguous -Quiet

    Assert ($res.Success -eq $true) 'Invoke-HeadlessMasterUpdater must succeed'
    Assert ($res.AcceptedRows -eq 5) "Expected 5 accepted rows (2 New + 3 Changed), got $($res.AcceptedRows)"
    Assert ($res.UpdatedCells -eq 4) "Expected 4 updated cells, got $($res.UpdatedCells)"
    Assert ($res.AddedRows -eq 2) "Expected 2 added rows, got $($res.AddedRows)"
    Assert (Test-Path $res.BackupPath) "Backup file must exist at $($res.BackupPath)"
    Assert (Test-Path $res.LogJsonlPath) "JSONL audit log must exist at $($res.LogJsonlPath)"
    Assert (Test-Path $res.LogTxtPath) "TXT audit log must exist at $($res.LogTxtPath)"

    # Verify updated value in base file
    $updatedBase = [FastExcelHelper]::ReadSheet($baseTest, 'Dzieci')
    $r6 = $updatedBase | Where-Object { $_.PSObject.Properties['IdDziecka'] -and $_.IdDziecka -eq 'D-1005' }
    Assert ($r6.TelefonRodzica -eq '600-999-888') "Expected updated phone 600-999-888, got $($r6.TelefonRodzica)"
}

# Test 4: Folder Watcher in -Once Mode (SWOT P4 / O1)
Test-Case 'Watch-IncomingFolder processes files in drop folder, moves to Archive, and generates report' {
    $dropDir = Join-Path $TestWorkspace 'WatchDrop'
    $archiveDir = Join-Path $dropDir 'Archive'
    $failedDir = Join-Path $dropDir 'Failed'
    $reportsDir = Join-Path $dropDir 'Reports'
    [void][System.IO.Directory]::CreateDirectory($dropDir)

    # Place incoming file into drop folder
    $incomingDropFile = Join-Path $dropDir 'monthly_incoming.xlsx'
    Copy-Item $incAOriginal $incomingDropFile

    $watcherScript = Join-Path $ProjectRoot 'Watch-IncomingFolder.ps1'
    $watchResults = & $watcherScript -FolderPath $dropDir -BaseFilePath $baseTest -Once -ExportHtmlReports -AutoAccept AllNonAmbiguous

    # Incoming file should be moved out of drop folder into Archive
    Assert (-not (Test-Path $incomingDropFile)) 'Processed incoming file must no longer be in drop folder'
    $archived = @(Get-ChildItem -Path $archiveDir -Filter "*monthly_incoming.xlsx")
    Assert ($archived.Count -eq 1) "Expected 1 archived file, found $($archived.Count)"

    # Report should be generated in Reports folder
    $reports = @(Get-ChildItem -Path $reportsDir -Filter "*monthly_incoming*.html")
    Assert ($reports.Count -eq 1) "Expected 1 HTML diff report in Reports folder, found $($reports.Count)"
}

# Test 5: Folder Watcher moves invalid/corrupt files to Failed/ folder
Test-Case 'Watch-IncomingFolder moves invalid file to Failed folder without crashing' {
    $dropDir = Join-Path $TestWorkspace 'WatchDrop'
    $failedDir = Join-Path $dropDir 'Failed'

    $corruptFile = Join-Path $dropDir 'corrupt_data.xlsx'
    [System.IO.File]::WriteAllBytes($corruptFile, [byte[]]@(0xDE, 0xAD, 0xBE, 0xEF))

    $watcherScript = Join-Path $ProjectRoot 'Watch-IncomingFolder.ps1'
    $watchResults = & $watcherScript -FolderPath $dropDir -BaseFilePath $baseTest -Once

    Assert (-not (Test-Path $corruptFile)) 'Corrupt file must be moved out of drop folder'
    $failedFiles = @(Get-ChildItem -Path $failedDir -Filter "*corrupt_data.xlsx")
    Assert ($failedFiles.Count -eq 1) "Expected 1 file in Failed folder, found $($failedFiles.Count)"
}

# Cleanup
try { Remove-Item -Recurse -Force $TestWorkspace -ErrorAction SilentlyContinue } catch { }

Write-Host ""
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Results: $($PassCount + $FailCount) total"
Write-Host " PASS: $PassCount" -ForegroundColor Green
if ($FailCount -gt 0) {
    Write-Host " FAIL: $FailCount" -ForegroundColor Red
    exit 1
} else {
    Write-Host " All Headless and Folder Watcher tests PASSED!" -ForegroundColor Green
}
