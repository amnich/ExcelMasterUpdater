#Requires -Version 5.1
<#
.SYNOPSIS
    Integration test suite verifying end-to-end synchronization on real sample files.

.DESCRIPTION
    Validates that Master-Updater accurately maps, compares, and writes back changes
    between SampleFiles\Baza_Uczniowie_Dowoz.xlsx and SampleFiles\Zmiany_Wrzesien_2026.xlsx,
    verifying diacritic handling, exact cell patching, row appending, backups, and idempotency.
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$sampleDir = Join-Path $ProjectRoot 'SampleFiles'

. $masterScript

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Real Sample Files Integration & Diacritics Verification" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

$pristineFixture = Join-Path (Join-Path $ProjectRoot 'TestFixtures') 'Baza_Uczniowie_Dowoz_Pristine.xlsx'
$baseSrc = if (Test-Path $pristineFixture) { $pristineFixture } else { Join-Path $sampleDir 'Baza_Uczniowie_Dowoz.xlsx' }
$incSrc  = Join-Path $sampleDir 'Zmiany_Wrzesien_2026.xlsx'

if (-not (Test-Path $baseSrc) -or -not (Test-Path $incSrc)) {
    throw "Sample files not found at $sampleDir"
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "MasterUpdater_SampleTest_$([System.Guid]::NewGuid().ToString('N').Substring(0, 8))"
[void][System.IO.Directory]::CreateDirectory($tempDir)

try {
    $baseCopy = Join-Path $tempDir 'Baza_Test.xlsx'
    Copy-Item $baseSrc $baseCopy

    $baseSheet = [FastExcelHelper]::GetSheetNames($baseCopy)[0]
    $incSheet  = [FastExcelHelper]::GetSheetNames($incSrc)[0]

    $baseHeaders = [FastExcelHelper]::GetHeaders($baseCopy, $baseSheet)
    $incHeaders  = [FastExcelHelper]::GetHeaders($incSrc, $incSheet)

    if ($baseHeaders.Count -lt 15 -or $incHeaders.Count -lt 10) {
        throw "Failed to extract headers with diacritics from sample files."
    }
    Write-Host "[PASS] Read $($baseHeaders.Count) base headers and $($incHeaders.Count) incoming headers with diacritics" -ForegroundColor Green

    $baseRaw = [FastExcelHelper]::ReadSheet($baseCopy, $baseSheet)
    $incRaw  = [FastExcelHelper]::ReadSheet($incSrc,  $incSheet)

    # 1. Auto-Map rules
    $mappingRules = [System.Collections.Generic.List[object]]::new()
    foreach ($ih in $incHeaders) {
        $normIh = $ih.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
        $matched = $null
        foreach ($bh in $baseHeaders) {
            $normBh = $bh.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
            if ($normBh -eq $normIh -or $normBh.Contains($normIh) -or $normIh.Contains($normBh)) {
                $matched = $bh
                break
            }
        }
        if ($matched) {
            $mappingRules.Add([PSCustomObject]@{
                BaseColumns   = @($matched)
                UpdateColumns = @($ih)
                MergeMode     = 'Exact'
                Separator     = ''
            })
        }
    }

    if ($mappingRules.Count -ne 10) {
        throw "Expected 10 auto-mapped rules, got $($mappingRules.Count)"
    }
    Write-Host "[PASS] 100% of incoming columns auto-mapped successfully (10/10)" -ForegroundColor Green

    # 2. Compare dataset
    $joinKey = @('Imię i Nazwisko Dziecka')
    $reviewItems = Invoke-MasterCompare -BaseRows $baseRaw `
                                        -IncomingRows $incRaw `
                                        -MappingRules $mappingRules `
                                        -BaseJoinKey $joinKey `
                                        -IncomingJoinKey $joinKey `
                                        -BaseHeaders $baseHeaders `
                                        -DetectRemoved $true

    $newRecs = @($reviewItems | Where-Object { $_.Status -eq 'New' })
    $chgRecs = @($reviewItems | Where-Object { $_.Status -eq 'Changed' })
    $remRecs = @($reviewItems | Where-Object { $_.Status -eq 'Removed' })
    $uncRecs = @($reviewItems | Where-Object { $_.Status -eq 'Unchanged' })

    if ($newRecs.Count -ne 2) { throw "Expected 2 new records, got $($newRecs.Count)" }
    if ($chgRecs.Count -ne 4) { throw "Expected 4 changed records, got $($chgRecs.Count)" }
    if ($uncRecs.Count -ne 10) { throw "Expected 10 unchanged records, got $($uncRecs.Count)" }
    if ($remRecs.Count -ne 6) { throw "Expected 6 removed records, got $($remRecs.Count)" }
    Write-Host "[PASS] Comparison accurately categorized: 2 New, 4 Changed, 10 Unchanged, 6 Removed" -ForegroundColor Green

    # 3. Mark New and Changed as Accepted and execute InPlace write-back
    foreach ($item in $reviewItems) {
        if ($item.Status -in @('New', 'Changed')) {
            $item.Decision = 'Accepted'
        }
    }

    $cfg = Get-DefaultAppConfig
    $cfg.BackupDirectory = Join-Path $tempDir 'Backups'
    $cfg.LogDirectory    = Join-Path $tempDir 'Logs'
    $cfg.WriteMode       = 'InPlace'

    $wbResult = Invoke-MasterWriteBack -BaseFilePath $baseCopy `
                                       -BaseSheet $baseSheet `
                                       -ReviewItems $reviewItems `
                                       -MappingRules $mappingRules `
                                       -AppConfig $cfg `
                                       -IncomingPath $incSrc `
                                       -IncomingSheet $incSheet

    if (-not $wbResult.Success) {
        throw "Writeback failed: $($wbResult.ErrorMessage)"
    }
    if ($wbResult.AddedRows -ne 2) {
        throw "Expected 2 added rows, got $($wbResult.AddedRows)"
    }
    if ($wbResult.UpdatedCells -ne 15) {
        throw "Expected 15 updated cells, got $($wbResult.UpdatedCells)"
    }
    if (-not (Test-Path $wbResult.BackupPath)) {
        throw "Backup file not found at $($wbResult.BackupPath)"
    }
    Write-Host "[PASS] InPlace writeback succeeded: 2 rows added, 15 cells patched, backup verified" -ForegroundColor Green

    # 4. Verify reloaded base file and assert idempotency
    $reloadedHeaders = [FastExcelHelper]::GetHeaders($baseCopy, $baseSheet)
    $reloadedRows    = [FastExcelHelper]::ReadSheet($baseCopy, $baseSheet)

    if ($reloadedRows.Count -ne 22) {
        throw "Expected 22 rows after append, got $($reloadedRows.Count)"
    }

    $secondCompare = Invoke-MasterCompare -BaseRows $reloadedRows `
                                          -IncomingRows $incRaw `
                                          -MappingRules $mappingRules `
                                          -BaseJoinKey $joinKey `
                                          -IncomingJoinKey $joinKey `
                                          -BaseHeaders $reloadedHeaders `
                                          -DetectRemoved $false

    $secNew = @($secondCompare | Where-Object { $_.Status -eq 'New' })
    $secChg = @($secondCompare | Where-Object { $_.Status -eq 'Changed' })
    $secUnc = @($secondCompare | Where-Object { $_.Status -eq 'Unchanged' })

    if ($secNew.Count -ne 0 -or $secChg.Count -ne 0 -or $secUnc.Count -ne 16) {
        throw "Idempotency assertion failed! New=$($secNew.Count), Changed=$($secChg.Count), Unchanged=$($secUnc.Count)"
    }
    Write-Host "[PASS] Re-import idempotency verified (0 New, 0 Changed, 16 Unchanged)" -ForegroundColor Green

    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host " ALL SAMPLE FILES INTEGRATION TESTS PASSED (100% SUCCESS)!" -ForegroundColor Green
    Write-Host "================================================================================" -ForegroundColor Green

} finally {
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
