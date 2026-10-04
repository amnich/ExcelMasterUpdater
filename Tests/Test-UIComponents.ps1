#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$masterScript = Join-Path $ProjectRoot 'Master-Updater.ps1'
$fixturesDir  = Join-Path $ProjectRoot 'TestFixtures'
$baseXlsx     = Join-Path $fixturesDir 'base.xlsx'
$incomingXlsx = Join-Path $fixturesDir 'incoming_A.xlsx'

Write-Host "Testing UI Components & Non-Interactive Launch..." -ForegroundColor Cyan

# 1. Test Dot-sourcing Master-Updater.ps1
. $masterScript
Write-Host "SUCCESS: Master-Updater.ps1 dot-sourced cleanly!" -ForegroundColor Green

# 2. Test Get-UiString across languages
foreach ($lang in @('pl', 'en', 'de')) {
    $script:CurrentLanguage = $lang
    $val = Get-UiString 'AppTitle'
    if ([string]::IsNullOrWhiteSpace($val)) {
        throw "Failed to get AppTitle for language $lang"
    }
    Write-Host "[$lang] AppTitle: $val" -ForegroundColor Gray
}
Write-Host "SUCCESS: Get-UiString verified across all 3 languages!" -ForegroundColor Green

# 3. Test Show-MasterUpdater in NonInteractive mode with parameters
[void](Show-MasterUpdater -InitBaseFilePath $baseXlsx -InitIncomingPath $incomingXlsx -InitBaseSheet 'Dzieci' -InitIncomingSheet 'AktualizacjaWrzesien' -NonInteractive)
Write-Host "SUCCESS: Show-MasterUpdater -NonInteractive initialized without exceptions!" -ForegroundColor Green

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "ALL UI COMPONENT TESTS PASSED!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
