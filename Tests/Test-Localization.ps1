#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$langPath = Join-Path $ProjectRoot 'language.json'

if (-not (Test-Path $langPath)) {
    throw "language.json not found at $langPath"
}

$raw = [System.IO.File]::ReadAllText($langPath, [System.Text.Encoding]::UTF8)
$data = $raw | ConvertFrom-Json

$languages = @('en', 'pl', 'de')
$keyMaps = @{}

foreach ($lang in $languages) {
    $prop = $data.Languages.PSObject.Properties[$lang]
    if (-not $prop) {
        throw "Language '$lang' missing from catalog!"
    }
    $stringsObj = $prop.Value.Strings
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($sp in $stringsObj.PSObject.Properties) {
        $keys.Add($sp.Name)
    }
    $keyMaps[$lang] = $keys
    Write-Host "[$lang] ($($prop.Value.DisplayName)): $($keys.Count) keys" -ForegroundColor Cyan
}

$masterKeys = $keyMaps['en'] | Sort-Object
$hasErrors = $false

foreach ($lang in @('pl', 'de')) {
    $currKeys = $keyMaps[$lang] | Sort-Object
    $missingInTarget = Compare-Object -ReferenceObject $masterKeys -DifferenceObject $currKeys | Where-Object { $_.SideIndicator -eq '<=' }
    $extraInTarget   = Compare-Object -ReferenceObject $masterKeys -DifferenceObject $currKeys | Where-Object { $_.SideIndicator -eq '=>' }

    if ($missingInTarget) {
        $hasErrors = $true
        Write-Error "Language '$lang' is missing keys: $($missingInTarget.InputObject -join ', ')"
    }
    if ($extraInTarget) {
        $hasErrors = $true
        Write-Error "Language '$lang' has extra keys: $($extraInTarget.InputObject -join ', ')"
    }
}

# Assert code-to-catalog parity against Master-Updater.ps1
$scriptFile = Join-Path $ProjectRoot 'Master-Updater.ps1'
if (Test-Path $scriptFile) {
    $scriptCode = [System.IO.File]::ReadAllText($scriptFile, [System.Text.Encoding]::UTF8)
    $regex = 'Get-UiString\s+[''"]([A-Za-z0-9_]+)[''"]'
    $matches = [System.Text.RegularExpressions.Regex]::Matches($scriptCode, $regex)
    $missingFromCatalog = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($m in $matches) {
        $k = $m.Groups[1].Value
        if ($masterKeys -notcontains $k) {
            [void]$missingFromCatalog.Add($k)
        }
    }
    if ($missingFromCatalog.Count -gt 0) {
        $hasErrors = $true
        Write-Error "Keys called in Master-Updater.ps1 but missing from language catalog ($($missingFromCatalog.Count)): $($missingFromCatalog -join ', ')"
    } else {
        Write-Host "SUCCESS: All unique Get-UiString keys in Master-Updater.ps1 ($($matches.Count) calls) are present in the catalog!" -ForegroundColor Green
    }
}

if (-not $hasErrors) {
    Write-Host "SUCCESS: 100% key completeness verified across all $($languages.Count) languages ($($masterKeys.Count) keys each)!" -ForegroundColor Green
} else {
    throw "Localization validation failed!"
}
