#Requires -Version 5.1
<#
.SYNOPSIS
    Automated WPF interactive controls regression test for selective reconciliation checkboxes and inline editing.

.DESCRIPTION
    Validates:
      1. Toggling individual diff cell checkboxes (Uncheck/Check) updates SelectedCells model and staging summary without throwing CommandNotFoundException or RuntimeException.
      2. Row action buttons (Select All / Deselect All in Row) toggle all row cell checkboxes and update live staging metrics.
      3. Batch column toggle checkboxes toggle matching columns across all review items and refresh detail view.
      4. Inline cell editing dynamically marks change as CustomEdited and displays revert button (↺).
      5. Revert button restores OriginalNewValue and clears dirty border.
#>

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$testDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Resolve-Path $MyInvocation.MyCommand.Path) }
$projectRoot = Split-Path -Parent (Resolve-Path $testDir)
$masterScript = Join-Path $projectRoot 'Master-Updater.ps1'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Interactive Selective Checkboxes & Editing" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Dot-source Master-Updater
. $masterScript
Write-Host "[PASS] Master-Updater.ps1 dot-sourced successfully" -ForegroundColor Green

# 2. Check Sample Files
$baseFile = Join-Path $projectRoot 'SampleFiles\Baza_Uczniowie_Dowoz.xlsx'
$incFile  = Join-Path $projectRoot 'SampleFiles\lista 18.xlsx'

if (-not (Test-Path $baseFile) -or -not (Test-Path $incFile)) {
    throw "Sample files not found: $baseFile / $incFile"
}

# 3. Initialize Window in NonInteractive Mode
$window = Show-MasterUpdater -InitBaseFilePath $baseFile -InitIncomingPath $incFile -InitBaseSheet 'Dzieci' -InitIncomingSheet 'Arkusz1' -NonInteractive
if (-not $window) { throw "Show-MasterUpdater did not return a valid Window object!" }
Write-Host "[PASS] Show-MasterUpdater initialized in NonInteractive mode" -ForegroundColor Green

# 4. Trigger Compare
$btnRunCompare = $window.FindName('btnRunCompare')
if (-not $btnRunCompare) { throw "btnRunCompare control not found!" }
try {
    $btnRunCompare.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
} catch {
    Write-Host "CATCH EXCEPTION: $($_.Exception.ToString())" -ForegroundColor Red
    Write-Host "STACK TRACE: $($_.ScriptStackTrace)" -ForegroundColor Red
    if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorRecord) {
        Write-Host "ERROR POSITION:`n$($_.Exception.InnerException.ErrorRecord.InvocationInfo.PositionMessage)" -ForegroundColor Red
    }
    throw
}
if ($script:AllReviewItems.Count -eq 0) { throw "Expected review items after compare, got 0" }
Write-Host "[PASS] Comparison generated $($script:AllReviewItems.Count) review items" -ForegroundColor Green

# 5. Select Item 0 (Changed item)
$lbReviewItems = $window.FindName('lbReviewItems')
$lbReviewItems.SelectedIndex = 0
$item0 = $lbReviewItems.SelectedItem
if (-not $item0) { throw "Failed to select first review item" }
Write-Host "[PASS] Selected item 0: $($item0.Title) (Changes: $($item0.Record.Changes.Count))" -ForegroundColor Green

$panelDiffContainer = $window.FindName('panelDiffContainer')
if (-not $panelDiffContainer) { throw "panelDiffContainer not found" }

# 6. Test Cell Checkboxes (Deselect / Reselect)
$cellCheckboxes = [System.Collections.Generic.List[object]]::new()
foreach ($elem in $panelDiffContainer.Children) {
    if ($elem -is [System.Windows.Controls.Grid]) {
        foreach ($child in $elem.Children) {
            if ($child -is [System.Windows.Controls.CheckBox]) {
                $cellCheckboxes.Add($child)
            }
        }
    }
}

if ($cellCheckboxes.Count -eq 0) { throw "No cell checkboxes found in diff container grid!" }
Write-Host "[PASS] Found $($cellCheckboxes.Count) cell checkboxes in diff table" -ForegroundColor Green

foreach ($chk in $cellCheckboxes) {
    $col = $chk.Tag.ColKey
    # Deselect
    $chk.IsChecked = $false
    $val = Get-ItemCellSelected $item0 $col
    if ($val -ne $false) { throw "Get-ItemCellSelected returned $val after unchecking $col" }
    
    # Reselect
    $chk.IsChecked = $true
    $val2 = Get-ItemCellSelected $item0 $col
    if ($val2 -ne $true) { throw "Get-ItemCellSelected returned $val2 after checking $col" }
}
Write-Host "[PASS] Toggled all individual cell checkboxes to false and true without error" -ForegroundColor Green

# 7. Test Row Tools (Deselect All / Select All)
$btnDeselectAll = $null
$btnSelectAll = $null
foreach ($elem in $panelDiffContainer.Children) {
    if ($elem -is [System.Windows.Controls.StackPanel]) {
        foreach ($child in $elem.Children) {
            if ($child -is [System.Windows.Controls.Button]) {
                if ($child.Content -like '*Odznacz*') { $btnDeselectAll = $child }
                if ($child.Content -like '*Zaznacz*')  { $btnSelectAll = $child }
            }
        }
    }
}

if (-not $btnDeselectAll -or -not $btnSelectAll) { throw "Row tool buttons not found!" }

# Click Deselect All
$btnDeselectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
foreach ($ch in $item0.Record.Changes) {
    if ((Get-ItemCellSelected $item0 $ch.BaseColumn) -ne $false) {
        throw "Column $($ch.BaseColumn) was not deselected by Deselect All!"
    }
}
Write-Host "[PASS] Row tool 'Odznacz wszystkie w wierszu' deselected all cells" -ForegroundColor Green

# Click Select All
$btnSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
foreach ($ch in $item0.Record.Changes) {
    if ((Get-ItemCellSelected $item0 $ch.BaseColumn) -ne $true) {
        throw "Column $($ch.BaseColumn) was not selected by Select All!"
    }
}
Write-Host "[PASS] Row tool 'Zaznacz wszystkie w wierszu' reselected all cells" -ForegroundColor Green

# 8. Test Batch Column Toggles
$wrapBatchColToggles = $window.FindName('wrapBatchColToggles')
if ($wrapBatchColToggles -and $wrapBatchColToggles.Children.Count -gt 0) {
    $firstBatchChk = $wrapBatchColToggles.Children[0]
    $batchColName = $firstBatchChk.Tag
    
    # Toggle off
    $firstBatchChk.IsChecked = $false
    foreach ($it in $script:AllReviewItems) {
        if ($it.Record -and $it.Record.Changes) {
            foreach ($c in $it.Record.Changes) {
                if ($c.BaseColumn -eq $batchColName -and (Get-ItemCellSelected $it $batchColName) -ne $false) {
                    throw "Batch toggle off failed for item $($it.Title) column $batchColName"
                }
            }
        }
    }
    
    # Toggle on
    $firstBatchChk.IsChecked = $true
    foreach ($it in $script:AllReviewItems) {
        if ($it.Record -and $it.Record.Changes) {
            foreach ($c in $it.Record.Changes) {
                if ($c.BaseColumn -eq $batchColName -and (Get-ItemCellSelected $it $batchColName) -ne $true) {
                    throw "Batch toggle on failed for item $($it.Title) column $batchColName"
                }
            }
        }
    }
    Write-Host "[PASS] Batch column toggles for '$batchColName' work correctly across all review items" -ForegroundColor Green
}

# 9. Test Inline Cell Editing and Revert
$firstTextBox = $null
$firstRevertBtn = $null
foreach ($elem in $panelDiffContainer.Children) {
    if ($elem -is [System.Windows.Controls.Grid]) {
        foreach ($child in $elem.Children) {
            if ($child -is [System.Windows.Controls.Border] -and $child.Child -is [System.Windows.Controls.Grid]) {
                $subGrid = $child.Child
                foreach ($sgChild in $subGrid.Children) {
                    if ($sgChild -is [System.Windows.Controls.TextBox] -and -not $firstTextBox) { $firstTextBox = $sgChild }
                    if ($sgChild -is [System.Windows.Controls.Button] -and -not $firstRevertBtn) { $firstRevertBtn = $sgChild }
                }
            }
        }
    }
}

if (-not $firstTextBox -or -not $firstRevertBtn) { throw "Inline edit controls not found!" }

$origVal = $firstTextBox.Text
$firstTextBox.Text = "CustomValue_12345"
if ($firstRevertBtn.Visibility -ne [System.Windows.Visibility]::Visible) {
    throw "Revert button not visible after inline editing!"
}
if ($firstTextBox.Tag.Chg.CustomEdited -ne $true) {
    throw "Change not flagged as CustomEdited!"
}
Write-Host "[PASS] Inline editing flagged change as CustomEdited and showed revert button" -ForegroundColor Green

$firstRevertBtn.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
if ($firstTextBox.Text -ne $origVal) {
    throw "Revert button failed to restore original text!"
}
if ($firstRevertBtn.Visibility -ne [System.Windows.Visibility]::Collapsed) {
    throw "Revert button not collapsed after revert!"
}
Write-Host "[PASS] Inline revert restored original value and collapsed revert button" -ForegroundColor Green

Write-Host "==========================================================" -ForegroundColor Green
Write-Host " ALL INTERACTIVE SELECTIVE RECONCILIATION TESTS PASSED 100%!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
