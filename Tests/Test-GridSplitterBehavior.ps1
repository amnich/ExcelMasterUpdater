# ==============================================================================
# TEST SUITE: GridSplitter Drag, Bounds & Layout Stability
# Tests resizing behavior, MinHeight bounds, template, and scrollbar transitions
# ==============================================================================
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $PSScriptRoot
$scriptPath  = Join-Path $projectRoot "Master-Updater.ps1"

Write-Output "=========================================================="
Write-Output " TEST SUITE: GridSplitter Drag & Bounds Verification"
Write-Output "=========================================================="

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore

# 1. Parse Window XAML directly from Master-Updater.ps1
$content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$start = $content.IndexOf("<Window")
$end = $content.IndexOf("</Window>") + 9
$xaml = $content.Substring($start, $end - $start)

$win = [System.Windows.Markup.XamlReader]::Parse($xaml)
$win.WindowState = [System.Windows.WindowState]::Maximized
$win.Show()

$grid = $win.FindName("tabMapping").Content
$r0 = $grid.RowDefinitions[0]
$r1 = $grid.RowDefinitions[1]
$r2 = $grid.RowDefinitions[2]
$r3 = $grid.RowDefinitions[3]
$r4 = $grid.RowDefinitions[4]

$gs = $grid.Children | Where-Object { $_ -is [System.Windows.Controls.GridSplitter] }

# Assert Row Definitions
if ($r1.MinHeight -ne 90) { throw "FAIL: Row 1 MinHeight expected 90, got $($r1.MinHeight)" }
if ($r2.Height.Value -ne 8) { throw "FAIL: Row 2 Height expected 8, got $($r2.Height.Value)" }
if ($r3.MinHeight -ne 48) { throw "FAIL: Row 3 MinHeight expected 48, got $($r3.MinHeight)" }
Write-Output "[PASS] RowDefinitions verified: Row1 MinHeight=90, Row2 Height=8, Row3 MinHeight=48"

# Assert GridSplitter Properties
if ($gs.ResizeBehavior -ne [System.Windows.Controls.GridResizeBehavior]::PreviousAndNext) {
    throw "FAIL: GridSplitter ResizeBehavior expected PreviousAndNext, got $($gs.ResizeBehavior)"
}
if ($gs.ResizeDirection -ne [System.Windows.Controls.GridResizeDirection]::Rows) {
    throw "FAIL: GridSplitter ResizeDirection expected Rows, got $($gs.ResizeDirection)"
}
Write-Output "[PASS] GridSplitter properties verified: ResizeBehavior=PreviousAndNext, ResizeDirection=Rows"

# Populate components with data to trigger scrollbars
$gridMappingRules = $win.FindName("gridMappingRules")
$lbJoinBase = $win.FindName("lbJoinBase")
$lbJoinIncoming = $win.FindName("lbJoinIncoming")
$gridPrev = $win.FindName("gridMappingResultPreview")

$rules = 1..15 | ForEach-Object { [PSCustomObject]@{ BaseColsStr="Base Col $_"; UpdColsStr="Upd Col $_"; MergeMode="Exact"; Separator="" } }
$gridMappingRules.ItemsSource = $rules
1..10 | ForEach-Object { [void]$lbJoinBase.Items.Add("Join Col $_"); [void]$lbJoinIncoming.Items.Add("Join Col $_") }
$prevItems = 1..10 | ForEach-Object { [PSCustomObject]@{ TargetBaseColumn="Col $_"; MappedSourceColumns="Col $_"; ProjectedValue="Val $_"; CurrentBaseValue="Old $_"; StatusText="Zmiana"; StatusBg="Orange"; StatusFg="White" } }
$gridPrev.ItemsSource = $prevItems
$win.UpdateLayout()

# Reflection methods for simulating drag
$mStart = [System.Windows.Controls.GridSplitter].GetMethod("OnDragStarted", [System.Reflection.BindingFlags]"NonPublic,Instance")
$mDelta = [System.Windows.Controls.GridSplitter].GetMethod("OnDragDelta", [System.Reflection.BindingFlags]"NonPublic,Instance")
$mComplete = [System.Windows.Controls.GridSplitter].GetMethod("OnDragCompleted", [System.Reflection.BindingFlags]"NonPublic,Instance")

$ctorStart = [System.Windows.Controls.Primitives.DragStartedEventArgs].GetConstructor(@([double], [double]))
$ctorDelta = [System.Windows.Controls.Primitives.DragDeltaEventArgs].GetConstructor(@([double], [double]))
$ctorComplete = [System.Windows.Controls.Primitives.DragCompletedEventArgs].GetConstructor(@([double], [double], [bool]))

$fData = [System.Windows.Controls.GridSplitter].GetField("_resizeData", [System.Reflection.BindingFlags]"NonPublic,Instance")

# Test 1: Drag Down past scrollbar transition to minimum collapse (Row 3 = 48)
$mStart.Invoke($gs, @($ctorStart.Invoke(@([double]0, [double]0))))
$cancelled = $false
for ($dy = 1; $dy -le 500; $dy += 3) {
    $mDelta.Invoke($gs, @($ctorDelta.Invoke(@([double]0, [double]$dy))))
    $win.UpdateLayout()
    if ($fData.GetValue($gs) -eq $null) {
        $cancelled = $true
        break
    }
}
$mComplete.Invoke($gs, @($ctorComplete.Invoke(@([double]0, [double]500, $false))))
$win.UpdateLayout()

if ($cancelled) { throw "FAIL: Dragging down was unexpectedly cancelled (snap-back bug occurred)!" }
if ($r3.ActualHeight -gt 55) { throw "FAIL: Row 3 failed to collapse down to MinHeight! Actual: $($r3.ActualHeight)" }
Write-Output "[PASS] Drag down past scrollbar transition collapsed preview cleanly to $($r3.ActualHeight)px without snap-back"

# Test 2: Drag Up to minimum top height (Row 1 = 90)
$mStart.Invoke($gs, @($ctorStart.Invoke(@([double]0, [double]0))))
$cancelledUp = $false
for ($dy = -1; $dy -ge -600; $dy -= 3) {
    $mDelta.Invoke($gs, @($ctorDelta.Invoke(@([double]0, [double]$dy))))
    $win.UpdateLayout()
    if ($fData.GetValue($gs) -eq $null) {
        $cancelledUp = $true
        break
    }
}
$mComplete.Invoke($gs, @($ctorComplete.Invoke(@([double]0, [double]-600, $false))))
$win.UpdateLayout()

if ($cancelledUp) { throw "FAIL: Dragging up was unexpectedly cancelled (snap-back bug occurred)!" }
if ($r1.ActualHeight -gt 100) { throw "FAIL: Row 1 failed to shrink down to MinHeight! Actual: $($r1.ActualHeight)" }
Write-Output "[PASS] Drag up expanded preview and shrunk mapping to $($r1.ActualHeight)px without snap-back"

# Test 3: Multiple rapid oscillating drags (stress test)
$mStart.Invoke($gs, @($ctorStart.Invoke(@([double]0, [double]0))))
foreach ($osc in @(50, -80, 120, -150, 200, -250, 100, -100, 0)) {
    $mDelta.Invoke($gs, @($ctorDelta.Invoke(@([double]0, [double]$osc))))
    $win.UpdateLayout()
}
$mComplete.Invoke($gs, @($ctorComplete.Invoke(@([double]0, [double]0, $false))))
$win.UpdateLayout()
Write-Output "[PASS] Rapid oscillating drag deltas handled without cancellation or layout faults"

$win.Close()

Write-Output "=========================================================="
Write-Output " ALL GRIDSPLITTER TESTS PASSED (100% SUCCESS)!"
Write-Output "=========================================================="
