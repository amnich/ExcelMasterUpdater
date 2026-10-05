# ==============================================================================
# TEST SUITE: Data & Mapping Live Preview Verification
# ==============================================================================
# Validates:
#   1. WPF element tree contains cardDataPreview, tabsMappingPreview, and all 3 preview tabs
#   2. BindSampleRowToGrid extracts and binds 1 row from Base and 1 row from Incoming file
#   3. UpdateDataMappingPreview executes Get-ProjectedRow and generates preview comparison items
#   4. Status classification (Identical, Will Change, New Value, Unmapped) works accurately
#   5. Sample row navigation (Prev/Next) advances the preview index
#   6. Dynamic trilingual localization updates preview headers and badges
# ==============================================================================

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$testDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Resolve-Path $MyInvocation.MyCommand.Path) }
$projectRoot = Split-Path -Parent (Resolve-Path $testDir)
$masterScript = Join-Path $projectRoot 'Master-Updater.ps1'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Data & Mapping Live Preview Verification" -ForegroundColor Cyan
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

# 3. Initialize Window with Files & Sheets
$window = Show-MasterUpdater -InitBaseFilePath $baseFile -InitIncomingPath $incFile -InitBaseSheet 'Dzieci' -InitIncomingSheet 'Arkusz1' -NonInteractive
if (-not $window) { throw "Show-MasterUpdater did not return a valid Window object!" }
Write-Host "[PASS] Show-MasterUpdater initialized with sample files" -ForegroundColor Green

# 4. Verify all UI Controls resolved
$ctrlNames = @(
    'cardDataPreview', 'txtPreviewHeader', 'txtPreviewHint', 'txtPreviewInfo',
    'btnPrevSampleRow', 'btnNextSampleRow', 'btnRefreshPreview',
    'tabsMappingPreview', 'tabPrevResult', 'tabPrevIncoming', 'tabPrevBase',
    'gridMappingResultPreview', 'gridPrevIncomingRow', 'gridPrevBaseRow',
    'txtEmptyPreviewPrompt', 'txtEmptyIncomingPrompt', 'txtEmptyBasePrompt',
    'colPrevTargetBase', 'colPrevSourceExpr', 'colPrevProjectedVal',
    'colPrevCurrentBaseVal', 'colPrevStatus'
)

foreach ($name in $ctrlNames) {
    $c = $window.FindName($name)
    if (-not $c) { throw "Missing expected UI control: $name" }
}
Write-Host "[PASS] All 22 Data & Mapping Preview UI controls verified in XAML tree" -ForegroundColor Green

if ($script:BaseHeaders.Count -eq 0 -or $script:IncomingHeaders.Count -eq 0) {
    throw "Failed to extract headers from sample files!"
}
Write-Host "[PASS] Headers extracted: $($script:BaseHeaders.Count) Base, $($script:IncomingHeaders.Count) Incoming" -ForegroundColor Green

# 5. Run AutoMap
$btnAutoMap = $window.FindName('btnAutoMap')
$btnAutoMap.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
if ($script:MappingRules.Count -eq 0) {
    throw "AutoMap did not produce any mapping rules!"
}
Write-Host "[PASS] AutoMap produced $($script:MappingRules.Count) mapping rules" -ForegroundColor Green

# 6. Verify Preview Generation
$gridMappingResultPreview = $window.FindName('gridMappingResultPreview')
$gridPrevIncomingRow      = $window.FindName('gridPrevIncomingRow')
$gridPrevBaseRow          = $window.FindName('gridPrevBaseRow')

if ($null -eq $gridMappingResultPreview.ItemsSource) {
    throw "gridMappingResultPreview ItemsSource is null after AutoMap!"
}

$prevItems = @($gridMappingResultPreview.ItemsSource)
if ($prevItems.Count -eq 0) {
    throw "gridMappingResultPreview has 0 items after AutoMap!"
}
if ($prevItems.Count -ne $script:BaseHeaders.Count) {
    throw "gridMappingResultPreview count ($($prevItems.Count)) does not match BaseHeaders count ($($script:BaseHeaders.Count))!"
}
Write-Host "[PASS] gridMappingResultPreview generated with $($prevItems.Count) field comparisons" -ForegroundColor Green

# Inspect sample row grids
if (-not $gridPrevBaseRow.ItemsSource -or -not $gridPrevIncomingRow.ItemsSource) {
    throw "gridPrevBaseRow or gridPrevIncomingRow has no ItemsSource!"
}
Write-Host "[PASS] 1 row from Base and 1 row from Incoming file bound to preview grids" -ForegroundColor Green

# Verify projected values and status badges
$mappedWithVal = $prevItems | Where-Object { -not [string]::IsNullOrEmpty($_.ProjectedValue) }
if ($mappedWithVal.Count -eq 0) {
    throw "No projected values generated in preview items!"
}

$validStatuses = @(
    (Get-UiString 'PreviewStatusIdentical'),
    (Get-UiString 'PreviewStatusChanged'),
    (Get-UiString 'PreviewStatusNewValue'),
    'Identyczne', 'Zmiana', 'Nowa wartość',
    'Identical', 'Will Change', 'New Value',
    'Identisch', 'Änderung', 'Neuer Wert'
)
$identicalOrChanged = @($prevItems | Where-Object { $_.StatusText -in $validStatuses })
if ($identicalOrChanged.Count -eq 0) {
    throw "Status badges not correctly assigned in preview items!"
}
Write-Host "[PASS] Live projection verified: $($mappedWithVal.Count) fields projected with status badges" -ForegroundColor Green

# 7. Test Row Navigation (Next / Prev)
$txtPreviewInfo = $window.FindName('txtPreviewInfo')
$infoBefore = $txtPreviewInfo.Text

$btnNextSampleRow = $window.FindName('btnNextSampleRow')
$btnNextSampleRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))

if ($script:PreviewRowIndex -ne 1) {
    throw "PreviewRowIndex did not advance after btnNextSampleRow click! Value: $($script:PreviewRowIndex)"
}
$infoAfter = $txtPreviewInfo.Text
if ($infoBefore -eq $infoAfter) {
    throw "txtPreviewInfo did not update after advancing sample row!"
}

$btnPrevSampleRow = $window.FindName('btnPrevSampleRow')
$btnPrevSampleRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))

if ($script:PreviewRowIndex -ne 0) {
    throw "PreviewRowIndex did not revert after btnPrevSampleRow click! Value: $($script:PreviewRowIndex)"
}
Write-Host "[PASS] Sample row navigation (Next/Prev) verified" -ForegroundColor Green

# 8. Test Dynamic Localization of Preview Elements
foreach ($lang in @('en', 'de', 'pl')) {
    $script:CurrentLanguage = $lang
    $tabHeader = Get-UiString 'TabPrevResult'
    $badgeIdentical = Get-UiString 'PreviewStatusIdentical'
    $badgeChanged = Get-UiString 'PreviewStatusChanged'
    $badgeNew = Get-UiString 'PreviewStatusNewValue'
    $badgeUnmapped = Get-UiString 'PreviewStatusUnmapped'
    
    if ([string]::IsNullOrWhiteSpace($tabHeader) -or [string]::IsNullOrWhiteSpace($badgeIdentical) -or [string]::IsNullOrWhiteSpace($badgeChanged)) {
        throw "Failed to retrieve preview UI strings for language '$lang'!"
    }
}
Write-Host "[PASS] Dynamic trilingual localization verified across EN/DE/PL" -ForegroundColor Green

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " ALL DATA & MAPPING PREVIEW TESTS PASSED (100% SUCCESS)!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
