#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'
$baseXlsx     = Join-Path $fixturesDir 'base.xlsx'
$incomingAXlsx = Join-Path $fixturesDir 'incoming_A.xlsx'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Show Unchanged Entries Option Verification" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Dot-source Master-Updater.ps1
. $masterScript
Write-Host "[PASS] Master-Updater.ps1 dot-sourced successfully" -ForegroundColor Green

# 2. Test Get-DefaultAppConfig and Get-AppConfig synchronization
$defCfg = Get-DefaultAppConfig
if (-not $defCfg.Contains('ShowUnchangedRows') -or $defCfg.ShowUnchangedRows -ne $false) {
    throw "DefaultAppConfig missing ShowUnchangedRows or not default `$false"
}
if (-not $defCfg.Contains('AutoSkipUnchanged') -or $defCfg.AutoSkipUnchanged -ne $true) {
    throw "DefaultAppConfig missing AutoSkipUnchanged or not default `$true"
}
Write-Host "[PASS] AppConfig defaults verified (ShowUnchangedRows=False, AutoSkipUnchanged=True)" -ForegroundColor Green

# 3. Test Get-StatusBadgeColors for Unchanged
foreach ($theme in @('Dark', 'Light')) {
    $b = Get-StatusBadgeColors -Status 'Unchanged' -ThemeName $theme
    if (-not $b.Bg -or -not $b.Fg) {
        throw "Get-StatusBadgeColors failed for Status 'Unchanged' in theme $theme"
    }
}
Write-Host "[PASS] Get-StatusBadgeColors returns valid palette for 'Unchanged'" -ForegroundColor Green

# 4. Initialize Show-MasterUpdater in NonInteractive mode with sample files
$window = Show-MasterUpdater -InitBaseFilePath $baseXlsx -InitIncomingPath $incomingAXlsx -InitBaseSheet 'Dzieci' -InitIncomingSheet 'AktualizacjaWrzesien' -NonInteractive

# Resolve controls from window
$chkShowUnchanged   = $window.FindName('chkShowUnchanged')
$cmbFilterStatus    = $window.FindName('cmbFilterStatus')
$cbiFilterUnchanged = $window.FindName('cbiFilterUnchanged')
$lbReviewItems      = $window.FindName('lbReviewItems')
$btnRunCompare      = $window.FindName('btnRunCompare')

$lbJoinBase         = $window.FindName('lbJoinBase')
$lbJoinIncoming     = $window.FindName('lbJoinIncoming')
$btnAutoMap         = $window.FindName('btnAutoMap')

if (-not $chkShowUnchanged)   { throw "Control chkShowUnchanged not found in XAML" }
if (-not $cbiFilterUnchanged) { throw "Control cbiFilterUnchanged not found in cmbFilterStatus" }

Write-Host "[PASS] UI controls chkShowUnchanged and cbiFilterUnchanged resolved from window" -ForegroundColor Green

# 5. Trigger comparison with default setting (ShowUnchanged = false)
$chkShowUnchanged.IsChecked = $false
$btnRunCompare.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))

$initialCount = $script:AllReviewItems.Count
if ($initialCount -ne 5) {
    throw "Expected 5 review items when unchanged are hidden, got $initialCount"
}
Write-Host "[PASS] Default compare without unchanged rows has 5 items (2 New, 3 Changed)" -ForegroundColor Green

# 6. Toggle chkShowUnchanged to true and verify review items re-population
$chkShowUnchanged.IsChecked = $true
$chkShowUnchanged.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))

$fullCount = $script:AllReviewItems.Count
if ($fullCount -ne 50) {
    throw "Expected 50 review items when unchanged rows are shown, got $fullCount"
}
$unchangedItems = @($script:AllReviewItems | Where-Object { $_.Record.Status -eq 'Unchanged' })
if ($unchangedItems.Count -ne 45) {
    throw "Expected 45 Unchanged items, got $($unchangedItems.Count)"
}
$sampleUnchanged = $unchangedItems[0]
if ($sampleUnchanged.Decision -ne 'Skipped') {
    throw "Expected default Decision for Unchanged items to be 'Skipped', got '$($sampleUnchanged.Decision)'"
}
if ([string]::IsNullOrWhiteSpace($sampleUnchanged.Subtitle)) {
    throw "Expected Subtitle on Unchanged item"
}
Write-Host "[PASS] Toggling 'Show unchanged' expanded list to 50 items (45 Unchanged with 'Skipped' decision)" -ForegroundColor Green

# 7. Test cmbFilterStatus filtering specifically to Unchanged rows
for ($i = 0; $i -lt $cmbFilterStatus.Items.Count; $i++) {
    $it = $cmbFilterStatus.Items[$i]
    if ($it.Tag -eq 'Unchanged') {
        $cmbFilterStatus.SelectedIndex = $i
        break
    }
}
if ($lbReviewItems.Items.Count -ne 45) {
    throw "Expected exactly 45 filtered items when cmbFilterStatus is 'Unchanged', got $($lbReviewItems.Items.Count)"
}
Write-Host "[PASS] cmbFilterStatus 'Unchanged' filter isolates all 45 unchanged records" -ForegroundColor Green

# 8. Test toggling chkShowUnchanged back to false reduces count back to 5
$chkShowUnchanged.IsChecked = $false
$chkShowUnchanged.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))

for ($i = 0; $i -lt $cmbFilterStatus.Items.Count; $i++) {
    $it = $cmbFilterStatus.Items[$i]
    if ($it.Tag -eq 'All') {
        $cmbFilterStatus.SelectedIndex = $i
        break
    }
}
if ($script:AllReviewItems.Count -ne 5) {
    throw "Expected count to revert to 5 when chkShowUnchanged is unchecked, got $($script:AllReviewItems.Count)"
}
Write-Host "[PASS] Unchecking 'Show unchanged' immediately drops unchanged items back to 5 items" -ForegroundColor Green

# 9. Verify WriteBack safety: Unchanged items must never patch cells if accepted
$dummyUnchanged = [PSCustomObject]@{
    Status         = 'Unchanged'
    Decision       = 'Accepted'
    Changes        = @()
    MatchedBaseRow = [PSCustomObject]@{ RowNumber = 10 }
    IncomingRow    = $null
}
$testRes = Invoke-MasterWriteBack -BaseFilePath $baseXlsx -BaseSheet 'Dzieci' -AcceptedItems @($dummyUnchanged) -AppConfig $script:AppConfig
if ($testRes.UpdatedCells -ne 0 -or $testRes.AddedRows -ne 0) {
    throw "Invoke-MasterWriteBack touched cells for unchanged item without modifications!"
}
Write-Host "[PASS] Invoke-MasterWriteBack guard verified (0 cells updated, 0 rows added for unchanged records)" -ForegroundColor Green

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "ALL SHOW-UNCHANGED TESTS PASSED (100% SUCCESS)!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
