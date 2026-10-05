#Requires -Version 5.1
<#
.SYNOPSIS
    Verification test suite for comparison mode options matching Compare-ExcelFiles:
    - Ignore case when comparing
    - Trim whitespace before comparing
    - Ignore punctuation & special characters (e.g. , . - / \)
    - Ignore internal whitespace differences when comparing fields
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'
$baseXlsx     = Join-Path $fixturesDir 'base.xlsx'
$incomingXlsx = Join-Path $fixturesDir 'incoming_A.xlsx'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " TEST SUITE: Comparison Mode Options Verification" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Dot-source Master-Updater.ps1
. $masterScript
Write-Host "[PASS] Master-Updater.ps1 dot-sourced successfully" -ForegroundColor Green

# 2. UI Controls Existence & Default Values
$win = Show-MasterUpdater -InitBaseFilePath $baseXlsx -InitIncomingPath $incomingXlsx -NonInteractive
if (-not $win) { throw "Show-MasterUpdater returned null window" }

$lblTitle = $win.FindName('lblCompareOptionsTitle')
$chkCase  = $win.FindName('chkIgnoreCase')
$chkTrim  = $win.FindName('chkTrimWhitespace')
$chkSpec  = $win.FindName('chkIgnoreSpecialChars')
$chkSpc   = $win.FindName('chkIgnoreAllSpaces')

if (-not $lblTitle) { throw "lblCompareOptionsTitle not found in XAML" }
if (-not $chkCase)  { throw "chkIgnoreCase not found in XAML" }
if (-not $chkTrim)  { throw "chkTrimWhitespace not found in XAML" }
if (-not $chkSpec)  { throw "chkIgnoreSpecialChars not found in XAML" }
if (-not $chkSpc)   { throw "chkIgnoreAllSpaces not found in XAML" }

if ($chkCase.IsChecked -ne $true) { throw "chkIgnoreCase default is not true" }
if ($chkTrim.IsChecked -ne $true) { throw "chkTrimWhitespace default is not true" }
if ($chkSpec.IsChecked -ne $true) { throw "chkIgnoreSpecialChars default is not true" }
if ($chkSpc.IsChecked -ne $true)  { throw "chkIgnoreAllSpaces default is not true" }
Write-Host "[PASS] All 4 comparison mode checkboxes resolved in UI and default to True" -ForegroundColor Green

# 3. Trilingual Localization of Options Title and Checkboxes
$keysToVerify = @(
    'LblCompareOptionsTitle',
    'LblIgnoreCase',
    'LblTrimWhitespace',
    'LblIgnoreSpecialChars',
    'LblIgnoreAllSpaces',
    'TooltipIgnoreCase',
    'TooltipTrimWhitespace',
    'TooltipIgnoreSpecialChars',
    'TooltipIgnoreAllSpaces'
)
foreach ($lang in @('pl', 'en', 'de')) {
    $script:CurrentLanguage = $lang
    foreach ($k in $keysToVerify) {
        $val = Get-UiString $k
        if ([string]::IsNullOrWhiteSpace($val)) {
            throw "[$lang] Key $k is missing or empty!"
        }
    }
}

$cmbLang = $win.FindName('cmbLanguage')
if ($cmbLang) {
    foreach ($cbi in $cmbLang.Items) {
        $cmbLang.SelectedItem = $cbi
        if ([string]::IsNullOrWhiteSpace($lblTitle.Text)) { throw "lblTitle is empty after selecting $($cbi.Tag)" }
        if ([string]::IsNullOrWhiteSpace($chkCase.Content)) { throw "chkCase is empty after selecting $($cbi.Tag)" }
    }
}
Write-Host "[PASS] Comparison options title, labels, and tooltips verified across PL, EN, DE" -ForegroundColor Green

# 4. Engine Test: IgnoreCase
$bRows = @( [PSCustomObject]@{ RowNumber = 2; Values = @{ 'ID' = '1'; 'City' = 'KRAKÓW' } } )
$iRows = @( [PSCustomObject]@{ SourceRowNumber = 2; Values = @{ 'ID' = '1'; 'City' = 'kraków' }; SourceFilePath = '' } )
$rules = @( [PSCustomObject]@{ BaseColumns = @('City'); UpdateColumns = @('City'); MergeMode = 'Exact'; Separator = '' } )

$resCaseIgnored = Invoke-MasterCompare -BaseRows $bRows -IncomingRows $iRows -MappingRules $rules -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $true; IgnoreAllSpaces = $true }
if ($resCaseIgnored[0].Changes.Count -ne 0) { throw "Expected 0 changes when IgnoreCase=true, got $($resCaseIgnored[0].Changes.Count)" }

$resCaseSensitive = Invoke-MasterCompare -BaseRows $bRows -IncomingRows $iRows -MappingRules $rules -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $false; Trim = $true; IgnoreSpecialChars = $true; IgnoreAllSpaces = $true }
if ($resCaseSensitive[0].Changes.Count -ne 1) { throw "Expected 1 change when IgnoreCase=false, got $($resCaseSensitive[0].Changes.Count)" }
Write-Host "[PASS] Engine: IgnoreCase option behaves correctly ($([bool]($resCaseIgnored[0].Changes.Count -eq 0)) vs $([bool]($resCaseSensitive[0].Changes.Count -eq 1)))" -ForegroundColor Green

# 5. Engine Test: TrimWhitespace
$bRowsTrim = @( [PSCustomObject]@{ RowNumber = 2; Values = @{ 'ID' = '1'; 'Name' = 'Anna Nowak' } } )
$iRowsTrim = @( [PSCustomObject]@{ SourceRowNumber = 2; Values = @{ 'ID' = '1'; 'Name' = '  Anna Nowak   ' }; SourceFilePath = '' } )
$rulesTrim = @( [PSCustomObject]@{ BaseColumns = @('Name'); UpdateColumns = @('Name'); MergeMode = 'Exact'; Separator = '' } )

$resTrimmed = Invoke-MasterCompare -BaseRows $bRowsTrim -IncomingRows $iRowsTrim -MappingRules $rulesTrim -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $false; IgnoreAllSpaces = $false }
if ($resTrimmed[0].Changes.Count -ne 0) { throw "Expected 0 changes when Trim=true, got $($resTrimmed[0].Changes.Count)" }

$resNotTrimmed = Invoke-MasterCompare -BaseRows $bRowsTrim -IncomingRows $iRowsTrim -MappingRules $rulesTrim -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $false; IgnoreSpecialChars = $false; IgnoreAllSpaces = $false }
if ($resNotTrimmed[0].Changes.Count -ne 1) { throw "Expected 1 change when Trim=false, got $($resNotTrimmed[0].Changes.Count)" }
Write-Host "[PASS] Engine: TrimWhitespace option behaves correctly" -ForegroundColor Green

# 6. Engine Test: IgnoreSpecialChars
$bRowsSpec = @( [PSCustomObject]@{ RowNumber = 2; Values = @{ 'ID' = '1'; 'Code' = '00-950' } } )
$iRowsSpec = @( [PSCustomObject]@{ SourceRowNumber = 2; Values = @{ 'ID' = '1'; 'Code' = '00950' }; SourceFilePath = '' } )
$rulesSpec = @( [PSCustomObject]@{ BaseColumns = @('Code'); UpdateColumns = @('Code'); MergeMode = 'Exact'; Separator = '' } )

$resSpecIgnored = Invoke-MasterCompare -BaseRows $bRowsSpec -IncomingRows $iRowsSpec -MappingRules $rulesSpec -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $true; IgnoreAllSpaces = $false }
if ($resSpecIgnored[0].Changes.Count -ne 0) { throw "Expected 0 changes when IgnoreSpecialChars=true, got $($resSpecIgnored[0].Changes.Count)" }

$resSpecDiff = Invoke-MasterCompare -BaseRows $bRowsSpec -IncomingRows $iRowsSpec -MappingRules $rulesSpec -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $false; IgnoreAllSpaces = $false }
if ($resSpecDiff[0].Changes.Count -ne 1) { throw "Expected 1 change when IgnoreSpecialChars=false, got $($resSpecDiff[0].Changes.Count)" }
Write-Host "[PASS] Engine: IgnoreSpecialChars option behaves correctly" -ForegroundColor Green

# 7. Engine Test: IgnoreAllSpaces
$bRowsSpc = @( [PSCustomObject]@{ RowNumber = 2; Values = @{ 'ID' = '1'; 'Address' = 'Ziśkoso 11 / 9' } } )
$iRowsSpc = @( [PSCustomObject]@{ SourceRowNumber = 2; Values = @{ 'ID' = '1'; 'Address' = 'Ziśkoso 11/9' }; SourceFilePath = '' } )
$rulesSpc = @( [PSCustomObject]@{ BaseColumns = @('Address'); UpdateColumns = @('Address'); MergeMode = 'Exact'; Separator = '' } )

$resSpcIgnored = Invoke-MasterCompare -BaseRows $bRowsSpc -IncomingRows $iRowsSpc -MappingRules $rulesSpc -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $true; IgnoreAllSpaces = $true }
if ($resSpcIgnored[0].Changes.Count -ne 0) { throw "Expected 0 changes when IgnoreAllSpaces=true, got $($resSpcIgnored[0].Changes.Count)" }

$resSpcDiff = Invoke-MasterCompare -BaseRows $bRowsSpc -IncomingRows $iRowsSpc -MappingRules $rulesSpc -BaseJoinKey @('ID') -IncomingJoinKey @('ID') -CompareOptions @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $false; IgnoreAllSpaces = $false }
if ($resSpcDiff[0].Changes.Count -ne 1) { throw "Expected 1 change when IgnoreAllSpaces=false, got $($resSpcDiff[0].Changes.Count)" }
Write-Host "[PASS] Engine: IgnoreAllSpaces option behaves correctly" -ForegroundColor Green

# 8. Profile Persistence of CompareOptions
$testProfDir = Join-Path ([System.IO.Path]::GetTempPath()) ("MU_ProfTest_" + [System.Guid]::NewGuid().ToString('N'))
try {
    $profToSave = [ordered]@{
        SchemaVersion     = '2.0'
        Name              = 'TestCompareProfile'
        CreatedUtc        = (Get-Date).ToUniversalTime().ToString('o')
        HeaderFingerprint = 'TEST_FP_1234'
        BaseFilePath      = $baseXlsx
        BaseSheet         = 'Dzieci'
        SourceSheetName   = 'Aktualizacja'
        JoinKeyBase       = @('ID')
        JoinKeyUpdate     = @('ID')
        MappingRules      = @()
        CompareOptions    = @{
            IgnoreCase         = $false
            TrimWhitespace     = $true
            IgnoreSpecialChars = $false
            IgnoreAllSpaces    = $true
        }
    }
    $savedPath = Save-MappingProfile -Profile $profToSave -StorePath $testProfDir
    if (-not (Test-Path $savedPath)) { throw "Failed to save profile to disk" }

    $loadedProf = Get-MappingProfileByFingerprint -Fingerprint 'TEST_FP_1234' -StorePath $testProfDir
    if (-not $loadedProf) { throw "Failed to load saved profile" }
    if ($loadedProf.CompareOptions.IgnoreCase -ne $false) { throw "Profile CompareOptions.IgnoreCase did not roundtrip" }
    if ($loadedProf.CompareOptions.TrimWhitespace -ne $true) { throw "Profile CompareOptions.TrimWhitespace did not roundtrip" }
    if ($loadedProf.CompareOptions.IgnoreSpecialChars -ne $false) { throw "Profile CompareOptions.IgnoreSpecialChars did not roundtrip" }
    if ($loadedProf.CompareOptions.IgnoreAllSpaces -ne $true) { throw "Profile CompareOptions.IgnoreAllSpaces did not roundtrip" }
    Write-Host "[PASS] ProfileStore correctly persists and restores CompareOptions" -ForegroundColor Green
} finally {
    if (Test-Path $testProfDir) { Remove-Item -Path $testProfDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "ALL COMPARISON MODE OPTION TESTS PASSED (100% SUCCESS)!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
