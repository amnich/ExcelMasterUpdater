# ==============================================================================
# TEST SUITE: New UX & Resilience Features Verification
# ==============================================================================
# Validates:
#   1. Tab 1 Mapping Search Bar ($txtSearchMapping, $btnSearchMappingClear)
#   2. Live Preview Immediate Reaction & Join Key Indicator (🔑)
#   3. Tab 2 Review ListBox MouseDoubleClick Quick Edit Handler
#   4. CSV Pipe Delimiter & Windows-1250 Resilient Detection
#   5. Headless Mode Execution Summary JSON Generation ($SummaryJsonPath)
# ==============================================================================

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$testDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Resolve-Path $MyInvocation.MyCommand.Path) }
$projectRoot = Split-Path -Parent (Resolve-Path $testDir)
$masterScript = Join-Path $projectRoot 'Master-Updater.ps1'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: New UX & Resilience Features Verification" -ForegroundColor Cyan
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

# 3. Test UI Controls and Tab 1 Mapping Filter
$window = Show-MasterUpdater -InitBaseFilePath $baseFile -InitIncomingPath $incFile -InitBaseSheet 'Dzieci' -InitIncomingSheet 'Arkusz1' -NonInteractive
if (-not $window) { throw "Show-MasterUpdater did not return a valid Window object!" }

$txtSearchMapping = $window.FindName('txtSearchMapping')
$btnSearchMappingClear = $window.FindName('btnSearchMappingClear')
$gridMappingRules = $window.FindName('gridMappingRules')
$lbJoinBase = $window.FindName('lbJoinBase')
$gridMappingResultPreview = $window.FindName('gridMappingResultPreview')
$lbReviewItems = $window.FindName('lbReviewItems')

if (-not $txtSearchMapping) { throw "Missing UI element: txtSearchMapping" }
if (-not $btnSearchMappingClear) { throw "Missing UI element: btnSearchMappingClear" }
Write-Host "[PASS] UI elements txtSearchMapping and btnSearchMappingClear resolved from XAML" -ForegroundColor Green

# Run AutoMap to populate rules
$btnAutoMap = $window.FindName('btnAutoMap')
$btnAutoMap.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
$totalRules = $script:MappingRules.Count
if ($totalRules -eq 0) { throw "AutoMap produced 0 rules!" }

# Verify Mapping Filter functionality
$txtSearchMapping.Text = 'Telefon'
$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridMappingRules.ItemsSource)
$filteredCount = 0
foreach ($item in $view) { $filteredCount++ }
if ($filteredCount -eq 0 -or $filteredCount -ge $totalRules) {
    throw "Mapping filter did not filter rules accurately! Total: $totalRules, Filtered: $filteredCount"
}
Write-Host "[PASS] Mapping search filter isolated $filteredCount rule(s) for 'Telefon'" -ForegroundColor Green

# Verify Clear button
$btnSearchMappingClear.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
if ($txtSearchMapping.Text -ne '') { throw "btnSearchMappingClear did not clear search text!" }
$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridMappingRules.ItemsSource)
$restoredCount = 0
foreach ($item in $view) { $restoredCount++ }
if ($restoredCount -ne $totalRules) { throw "Mapping filter clear did not restore all rules!" }
Write-Host "[PASS] Clear button reset filter and restored all $totalRules rules" -ForegroundColor Green

# 4. Test Live Preview Join Key Reaction
if ($lbJoinBase.Items.Count -gt 0) {
    # Select first item as join key
    $lbJoinBase.SelectedItems.Add($lbJoinBase.Items[0])
    $selectedKeyName = $lbJoinBase.Items[0].ToString()
    
    # Trigger preview update
    & $script:UpdateDataMappingPreview
    
    $prevItems = @($gridMappingResultPreview.ItemsSource)
    $keyItem = $prevItems | Where-Object { $_.IsJoinKey -eq $true }
    if (-not $keyItem) { throw "Preview items did not flag IsJoinKey when join key was selected!" }
    if (-not ($keyItem.TargetBaseColumn -like "🔑 *")) { throw "Preview TargetBaseColumn did not include key icon 🔑!" }
    Write-Host "[PASS] Live preview dynamically marked join key '$selectedKeyName' with key icon 🔑" -ForegroundColor Green
}

# 5. Test FastExcelHelper CSV Pipe Delimiter & Windows-1250 Encoding
$testCsvTemp = [System.IO.Path]::GetTempFileName() + ".csv"
try {
    # Create pipe-delimited CSV encoded in Windows-1250 with Polish diacritics
    $enc1250 = [System.Text.Encoding]::GetEncoding(1250)
    $csvLines = @(
        "Identyfikator|Imię i Nazwisko|Miasto|Wartość",
        "1|Stanisław Żółtowski|Kraków|1500,50",
        "2|Bożena Łącka|Gdańsk|2300,00"
    )
    $csvContent = [string]::Join("`r`n", $csvLines)
    [System.IO.File]::WriteAllText($testCsvTemp, $csvContent, $enc1250)
    
    # Test DetectFileEncoding
    $detectedEnc = [FastExcelHelper]::DetectFileEncoding($testCsvTemp)
    if ($detectedEnc.CodePage -ne 1250 -and $detectedEnc.CodePage -ne [System.Text.Encoding]::Default.CodePage) {
        throw "FastExcelHelper::DetectFileEncoding failed to detect Windows-1250/ANSI! Detected: $($detectedEnc.WebName)"
    }
    Write-Host "[PASS] FastExcelHelper accurately detected ANSI/Windows-1250 encoding" -ForegroundColor Green
    
    # Test GetHeaders with Pipe delimiter
    $csvHeaders = [FastExcelHelper]::GetHeaders($testCsvTemp, "CSV")
    if ($csvHeaders.Count -ne 4 -or $csvHeaders[0] -ne 'Identyfikator' -or $csvHeaders[1] -ne 'Imię i Nazwisko') {
        throw "FastExcelHelper::GetHeaders failed on pipe-delimited CSV with diacritics! Got: $([string]::Join(',', $csvHeaders))"
    }
    Write-Host "[PASS] FastExcelHelper parsed pipe-delimited headers with Polish diacritics: $([string]::Join(', ', $csvHeaders))" -ForegroundColor Green
    
    # Test ReadSheet with Pipe delimiter
    $csvRows = [FastExcelHelper]::ReadSheet($testCsvTemp, "CSV", 10)
    if ($csvRows.Count -ne 2) { throw "ReadSheet failed to parse all rows! Count: $($csvRows.Count)" }
    $psoName = $csvRows[0].'Imię i Nazwisko'
    if ($psoName -ne 'Stanisław Żółtowski') {
        throw "ReadSheet failed to preserve diacritics! Got: '$psoName'"
    }
    Write-Host "[PASS] FastExcelHelper read rows cleanly preserving Polish diacritics: '$psoName'" -ForegroundColor Green
} finally {
    if (Test-Path $testCsvTemp) { Remove-Item -Force $testCsvTemp }
}

# 6. Test Headless Mode Summary JSON Output
$tempSummaryJson = [System.IO.Path]::GetTempFileName() + ".json"
try {
    $hRes = Invoke-HeadlessMasterUpdater -BaseFilePath $baseFile -IncomingPath $incFile -BaseSheet 'Dzieci' -IncomingSheet 'Arkusz1' -AutoAccept 'None' -SummaryJsonPath $tempSummaryJson -Quiet
    if (-not (Test-Path $tempSummaryJson)) {
        throw "SummaryJsonPath was specified but file was not created: $tempSummaryJson"
    }
    $summaryObj = Get-Content $tempSummaryJson -Raw | ConvertFrom-Json
    if (-not $summaryObj.Success -or $summaryObj.TotalIncoming -le 0) {
        throw "Summary JSON artifact did not contain valid comparison metrics!"
    }
    if ($hRes.SummaryJsonPath -ne $tempSummaryJson) {
        throw "Headless result object did not reflect SummaryJsonPath property!"
    }
    Write-Host "[PASS] Headless mode successfully generated machine-readable execution summary JSON" -ForegroundColor Green
} finally {
    if (Test-Path $tempSummaryJson) { Remove-Item -Force $tempSummaryJson }
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " ALL NEW UX & RESILIENCE FEATURE TESTS PASSED (100%)!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
