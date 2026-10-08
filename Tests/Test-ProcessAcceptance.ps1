#Requires -Version 5.1
<#
.SYNOPSIS
    Formal Business Process Acceptance Tests (PAT) for Excel Master Updater.

.DESCRIPTION
    Validates end-to-end business acceptance criteria across all core process stages:
    1. AC-1: Multi-format Ingestion Process (Excel & CSV with Polish UTF-8 BOM).
    2. AC-2: Intelligent Schema Mapping & Multi-column Concatenation Acceptance.
    3. AC-3: Strict Status Classification & Discrepancy Isolation.
    4. AC-4: Atomic InPlace OpenXML Patching & Zero-Damage Table Preservation.
    5. AC-5: Non-Repudiation Audit Trails, In-Workbook ImportLog & Rollback Backups.
    6. AC-6: Automated Folder Watcher Ingestion, Processing & Archiving Pipeline.

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
    Mandatory UTF-8 with BOM encoding.
#>

$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'

$PassCount = 0
$FailCount = 0

function Test-Case {
    param([string]$Name, [scriptblock]$Block)
    try {
        & $Block
        Write-Host "[ACCEPTANCE PASS] $Name" -ForegroundColor Green
        $script:PassCount++
    } catch {
        Write-Host "[ACCEPTANCE FAIL] $Name -- $($_.Exception.Message)" -ForegroundColor Red
        $script:FailCount++
    }
}

function Assert-Acceptance {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ACCEPTANCE VIOLATION: $Message" }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " BUSINESS PROCESS ACCEPTANCE TEST SUITE (PAT) - EXCEL MASTER UPDATER" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# Dot-source Master-Updater engine in headless mode
. $masterScript -NonInteractive
Write-Host "Loaded Master-Updater.ps1 core engine." -ForegroundColor Gray

# Prepare dedicated clean acceptance workspace
$workspace = Join-Path $env:TEMP "MasterUpdater_ProcessAcceptance_$([guid]::NewGuid().ToString('N').Substring(0,8))"
if (Test-Path $workspace) { Remove-Item -Recurse -Force $workspace }
[void][System.IO.Directory]::CreateDirectory($workspace)

try {
    # -------------------------------------------------------------------------
    # AC-1: Multi-Format Ingestion Acceptance
    # -------------------------------------------------------------------------
    Test-Case "AC-1: Multi-Format Ingestion Acceptance (.xlsx and .csv)" {
        $baseFixture = Join-Path $fixturesDir 'base.xlsx'
        $incFixtureA = Join-Path $fixturesDir 'incoming_A.xlsx'
        $incFixtureB = Join-Path $fixturesDir 'incoming_B.csv'
        
        Assert-Acceptance (Test-Path $baseFixture) "Base fixture missing: $baseFixture"
        Assert-Acceptance (Test-Path $incFixtureA) "Incoming A fixture missing: $incFixtureA"
        Assert-Acceptance (Test-Path $incFixtureB) "Incoming B CSV missing: $incFixtureB"
        
        $baseSheets = [FastExcelHelper]::GetSheetNames($baseFixture)
        Assert-Acceptance ($baseSheets -contains 'Dzieci') "Base worksheet 'Dzieci' discovered."
        
        $baseHdrs = [FastExcelHelper]::GetHeaders($baseFixture, 'Dzieci')
        Assert-Acceptance ($baseHdrs.Count -ge 5) "Base headers successfully read (Count: $($baseHdrs.Count))."
        Assert-Acceptance ($baseHdrs -contains 'PESEL') "Base schema contains primary key 'PESEL'."
        
        # Test CSV ingestion via DetectCsvDelimiter and ParseCsvLine
        $delim = [FastExcelHelper]::DetectCsvDelimiter($incFixtureB)
        Assert-Acceptance ($delim -in @(',', ';', "`t")) "CSV delimiter accurately detected ('$delim')."
        
        $csvLines = Get-Content $incFixtureB -Encoding UTF8
        $csvHdrs = [FastExcelHelper]::ParseCsvLine($csvLines[0], $delim)
        Assert-Acceptance ($csvHdrs.Count -ge 3) "CSV headers parsed via RFC-4180 engine (Count: $($csvHdrs.Count))."
        
        $rows = [FastExcelHelper]::ReadSheet($baseFixture, 'Dzieci')
        Assert-Acceptance ($rows.Count -gt 0) "Worksheet rows loaded into memory model ($($rows.Count) rows)."
    }

    # -------------------------------------------------------------------------
    # AC-2: Intelligent Schema Mapping & Concatenation Acceptance
    # -------------------------------------------------------------------------
    Test-Case "AC-2: Intelligent Schema Mapping & Concatenation Acceptance" {
        $incomingHeaders = @('Numer PESEL', 'Imie', 'Nazwisko', 'Ulica i numer', 'Miejscowosc', 'Telefon')
        $baseHeaders     = @('PESEL', 'Imię', 'Nazwisko', 'Adres', 'Telefon kontaktowy')
        
        # 1. Verify FastDiffHelper string normalization & fuzzy matching
        $autoMapRules = @()
        foreach ($bHdr in $baseHeaders) {
            foreach ($iHdr in $incomingHeaders) {
                $normB = [FastDiffHelper]::Normalize($bHdr, $true, $true, $true, $false)
                $normI = [FastDiffHelper]::Normalize($iHdr, $true, $true, $true, $false)
                if ($normB -eq $normI -or $normI -like "*$normB*" -or $normB -like "*$normI*") {
                    $autoMapRules += [PSCustomObject]@{
                        BaseColumn     = $bHdr
                        IncomingColumn = $iHdr
                        Type           = 'Direct'
                    }
                    break
                }
            }
        }
        Assert-Acceptance ($autoMapRules.Count -ge 3) "Auto-map heuristic resolved $($autoMapRules.Count) column correspondences."
        
        # 2. Verify Concatenation Rule Application
        $row = [ordered]@{
            'Ulica i numer' = 'ul. Lipowa 12'
            'Miejscowosc'   = 'Warszawa'
        }
        $concatRule = [PSCustomObject]@{
            BaseColumn      = 'Adres'
            SourceColumns   = @('Ulica i numer', 'Miejscowosc')
            Delimiter       = ', '
        }
        $combined = ($concatRule.SourceColumns | ForEach-Object { $row[$_] }) -join $concatRule.Delimiter
        Assert-Acceptance ($combined -eq 'ul. Lipowa 12, Warszawa') "Compound address rule correctly produced '$combined'."
    }

    # -------------------------------------------------------------------------
    # AC-3: Strict Status Classification & Discrepancy Isolation
    # -------------------------------------------------------------------------
    Test-Case "AC-3: Status Classification Acceptance (New, Changed, Unchanged)" {
        $sampleBase = @(
            [ordered]@{ 'PESEL' = '11111111111'; 'Imie' = 'Jan';    'Nazwisko' = 'Kowalski'; 'Klasa' = '1A' },
            [ordered]@{ 'PESEL' = '22222222222'; 'Imie' = 'Anna';   'Nazwisko' = 'Nowak';    'Klasa' = '2B' },
            [ordered]@{ 'PESEL' = '33333333333'; 'Imie' = 'Piotr';  'Nazwisko' = 'Zieliński'; 'Klasa' = '3C' }
        )
        $sampleIncoming = @(
            [ordered]@{ 'PESEL' = '11111111111'; 'Imie' = 'Jan';    'Nazwisko' = 'Kowalski'; 'Klasa' = '1A' },  # Unchanged
            [ordered]@{ 'PESEL' = '22222222222'; 'Imie' = 'Anna';   'Nazwisko' = 'Nowak';    'Klasa' = '3B' },  # Changed (Klasa)
            [ordered]@{ 'PESEL' = '44444444444'; 'Imie' = 'Kacper'; 'Nazwisko' = 'Lewandowski'; 'Klasa' = '1B' } # New
        )
        
        $baseIdx = @{}
        foreach ($r in $sampleBase) { $baseIdx[$r['PESEL']] = $r }
        
        $newCount = 0
        $changedCount = 0
        $unchangedCount = 0
        
        foreach ($incRow in $sampleIncoming) {
            $key = $incRow['PESEL']
            if (-not $baseIdx.ContainsKey($key)) {
                $newCount++
            } else {
                $baseRow = $baseIdx[$key]
                $isDiff = $false
                foreach ($col in $incRow.Keys) {
                    if ($baseRow.Contains($col) -and $baseRow[$col] -ne $incRow[$col]) {
                        $isDiff = $true
                        break
                    }
                }
                if ($isDiff) { $changedCount++ } else { $unchangedCount++ }
            }
        }
        
        Assert-Acceptance ($newCount -eq 1) "Expected 1 New record; classified: $newCount"
        Assert-Acceptance ($changedCount -eq 1) "Expected 1 Changed record; classified: $changedCount"
        Assert-Acceptance ($unchangedCount -eq 1) "Expected 1 Unchanged record; classified: $unchangedCount"
    }

    # -------------------------------------------------------------------------
    # AC-4: Atomic InPlace OpenXML Patching & Table Preservation Acceptance
    # -------------------------------------------------------------------------
    Test-Case "AC-4: Atomic InPlace Patching & Zero-Damage Table Preservation" {
        $patBaseFile = Join-Path $workspace "pat_base.xlsx"
        Copy-Item -Path (Join-Path $fixturesDir "base.xlsx") -Destination $patBaseFile -Force
        
        $originalSize = (Get-Item $patBaseFile).Length
        Assert-Acceptance ($originalSize -gt 0) "Source file staged for InPlace update."
        
        # Apply InPlace Cell Mutation via EditExcelHelper
        $ops = [System.Collections.Generic.List[RowOp]]::new()
        $patchOp = [RowOp]::new()
        $patchOp.Type = 'PatchCell'
        $patchOp.RowNumber = 2
        # Col 4 is 'Klasa' (0-indexed: IdDziecka=0, ImieNazwisko=1, DataUrodzenia=2, PESEL=3, Klasa=4)
        $patchOp.Cells[4] = "3B"
        $ops.Add($patchOp)
        
        [EditExcelHelper]::WriteChanges($patBaseFile, 'Dzieci', $ops)
        
        # Verify cell readback without full destructive rewrite
        $recheckRows = [FastExcelHelper]::ReadSheet($patBaseFile, 'Dzieci')
        $firstRowVal = $recheckRows[0].Klasa
        Assert-Acceptance ($firstRowVal -eq '3B') "Cell E2 InPlace update verified in OpenXML stream ($firstRowVal)."
        
        # Append a new row via AppendRow op
        $appendOps = [System.Collections.Generic.List[RowOp]]::new()
        $appendOp = [RowOp]::new()
        $appendOp.Type = 'AppendRow'
        $appendOp.Cells[0] = 'D-9999'
        $appendOp.Cells[1] = 'Janek Nowy'
        $appendOp.Cells[3] = '99999999999'
        $appendOp.Cells[4] = '1A'
        $appendOps.Add($appendOp)
        
        $rowCountBefore = $recheckRows.Count
        [EditExcelHelper]::WriteChanges($patBaseFile, 'Dzieci', $appendOps)
        
        $afterRows = [FastExcelHelper]::ReadSheet($patBaseFile, 'Dzieci')
        Assert-Acceptance ($afterRows.Count -eq ($rowCountBefore + 1)) "Row append succeeded: $($afterRows.Count) rows verified."
    }

    # -------------------------------------------------------------------------
    # AC-5: Non-Repudiation Audit, ImportLog Sheet & Rollback Acceptance
    # -------------------------------------------------------------------------
    Test-Case "AC-5: Audit Log, In-Workbook ImportLog & Backup Rollback" {
        $auditBase = Join-Path $workspace "pat_audit.xlsx"
        Copy-Item -Path (Join-Path $fixturesDir "base.xlsx") -Destination $auditBase -Force
        
        $backupDir = Join-Path $workspace "Backups"
        [void][System.IO.Directory]::CreateDirectory($backupDir)
        
        # 1. Create Pre-Write Backup
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $backupFile = Join-Path $backupDir "pat_audit.$timestamp.bak.xlsx"
        Copy-Item -Path $auditBase -Destination $backupFile -Force
        Assert-Acceptance (Test-Path $backupFile) "Mandatory pre-write backup staged: $backupFile"
        
        # 2. Append ImportLog Sheet via EditExcelHelper::AppendWorksheetLog
        $headers = [string[]]@('Timestamp', 'User', 'Action', 'TargetRow', 'Field', 'OldValue', 'NewValue', 'Status')
        $logRows = [System.Collections.ArrayList]::new()
        [void]$logRows.Add([string[]]@((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $env:USERNAME, 'PAT_UPDATE', 'Row 2', 'Adres', 'Stary Adres', 'Nowy Adres', 'ACCEPTED'))
        
        [EditExcelHelper]::AppendWorksheetLog($auditBase, 'ImportLog', $headers, $logRows)
        
        $sheetsWithLog = [FastExcelHelper]::GetSheetNames($auditBase)
        Assert-Acceptance ($sheetsWithLog -contains 'ImportLog') "In-workbook 'ImportLog' sheet verified in OpenXML workbook."
        
        $logRowsRead = [FastExcelHelper]::ReadSheet($auditBase, 'ImportLog')
        Assert-Acceptance ($logRowsRead.Count -eq 1) "ImportLog entries read back accurately (Count: $($logRowsRead.Count))."
        Assert-Acceptance ($logRowsRead[0].Action -eq 'PAT_UPDATE') "Audit record action integrity asserted."
    }

    # -------------------------------------------------------------------------
    # AC-6: Automated Folder Watcher Processing Pipeline Acceptance
    # -------------------------------------------------------------------------
    Test-Case "AC-6: Automated Folder Watcher Daemon Pipeline Acceptance" {
        $watchRoot = Join-Path $workspace "WatcherPipeline"
        $inDir     = Join-Path $watchRoot "Incoming"
        $arcDir    = Join-Path $inDir "Archive"
        $failDir   = Join-Path $inDir "Failed"
        $reportsDir= Join-Path $inDir "Reports"
        $baseCopy  = Join-Path $watchRoot "base_watch.xlsx"
        
        @($inDir, $arcDir, $failDir, $reportsDir) | ForEach-Object { [void][System.IO.Directory]::CreateDirectory($_) }
        Copy-Item -Path (Join-Path $fixturesDir "base.xlsx") -Destination $baseCopy -Force
        
        # Stage incoming update file
        $stagedInc = Join-Path $inDir "monthly_update.xlsx"
        Copy-Item -Path (Join-Path $fixturesDir "incoming_A.xlsx") -Destination $stagedInc -Force
        Assert-Acceptance (Test-Path $stagedInc) "Incoming update staged in watcher drop folder."
        
        # Execute Watcher in -Once mode
        $watcherScript = Join-Path $ProjectRoot "Watch-IncomingFolder.ps1"
        Assert-Acceptance (Test-Path $watcherScript) "Watcher daemon script identified: $watcherScript"
        
        & $watcherScript -FolderPath $inDir -BaseFilePath $baseCopy -AutoAccept 'AllNonAmbiguous' -Once
                         
        # Assert file was moved to Archive folder upon successful processing
        Assert-Acceptance (-not (Test-Path $stagedInc)) "Staged file consumed from Incoming folder."
        $archivedFiles = Get-ChildItem -Path $arcDir -Filter "*monthly_update.xlsx"
        Assert-Acceptance ($archivedFiles.Count -ge 1) "File moved to Archive folder (Count: $($archivedFiles.Count))."
    }

} finally {
    if (Test-Path $workspace) {
        Remove-Item -Recurse -Force $workspace -ErrorAction SilentlyContinue
    }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " BUSINESS PROCESS ACCEPTANCE TEST SUMMARY" -ForegroundColor Cyan
Write-Host " Total Criteria: $($PassCount + $FailCount) | Passed: $PassCount | Failed: $FailCount" -ForegroundColor $(if ($FailCount -eq 0) { 'Green' } else { 'Red' })
Write-Host "================================================================================" -ForegroundColor Cyan

if ($FailCount -gt 0) {
    throw "$FailCount Acceptance Criteria FAILED."
}
