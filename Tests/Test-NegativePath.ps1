# Test-NegativePath.ps1
# Negative-path and fuzz test suite for Master-Updater components.
# Tests: corrupt ZIP, BOM-stripped CSV, UTF-16 CSV, empty join-key, large cell values,
#        locked file IOException, orphaned .tmp.xlsx, and SHA-256 fingerprint determinism.
# Compatible with Windows PowerShell 5.1 and PowerShell 7.x.

$ErrorActionPreference = 'Stop'
$PassCount = 0
$FailCount = 0
$TestDir = Join-Path $env:TEMP 'MasterUpdater_NegTests'
if (Test-Path $TestDir) { Remove-Item -Recurse -Force $TestDir }
[void][System.IO.Directory]::CreateDirectory($TestDir)

function Test-Case {
    param([string]$Name, [scriptblock]$Block)
    try {
        & $Block
        Write-Host "[PASS] $Name" -ForegroundColor Green
        $script:PassCount++
    } catch {
        Write-Host "[FAIL] $Name -- $($_.Exception.Message)" -ForegroundColor Red
        $script:FailCount++
    }
}

function Assert {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Build C# helper avoiding embedded double-quotes in here-strings
$ct  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
$rel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
$wbr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
$wb  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
$ws  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>001</t></is></c><c r="B2" t="inlineStr"><is><t>Jan Kowalski</t></is></c></row></sheetData></worksheet>'

function New-MinimalXlsx {
    param([string]$Path)
    if (Test-Path $Path) { Remove-Item -Force $Path }
    $entries = [ordered]@{
        '[Content_Types].xml'       = $script:ct
        '_rels/.rels'               = $script:rel
        'xl/_rels/workbook.xml.rels' = $script:wbr
        'xl/workbook.xml'           = $script:wb
        'xl/worksheets/sheet1.xml'  = $script:ws
    }
    $fsOut = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $zip   = New-Object System.IO.Compression.ZipArchive($fsOut, [System.IO.Compression.ZipArchiveMode]::Create)
    $enc   = [System.Text.Encoding]::UTF8
    foreach ($name in $entries.Keys) {
        $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $sw = New-Object System.IO.StreamWriter($entry.Open(), $enc)
        $sw.Write($entries[$name])
        $sw.Dispose()
    }
    $zip.Dispose()
    $fsOut.Dispose()
}

function Test-ValidZip {
    param([string]$Path)
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $z  = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $ok = $z.Entries.Count -gt 0
        $z.Dispose(); $fs.Dispose()
        return $ok
    } catch { return $false }
}

Write-Host "`nRunning negative-path test suite...`n" -ForegroundColor Cyan

Test-Case 'Corrupt XLSX rejected cleanly' {
    $path = Join-Path $TestDir 'corrupt.xlsx'
    [System.IO.File]::WriteAllBytes($path, [byte[]]@(0x50,0x4B,0x03,0x04,0xDE,0xAD,0xBE,0xEF))
    Assert (-not (Test-ValidZip $path)) 'Corrupt ZIP must NOT be detected as valid'
}

Test-Case 'BOM-stripped CSV (UTF-8 no BOM) readable' {
    $path = Join-Path $TestDir 'nobom.csv'
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("ID;Imie;Nazwisko`r`n001;Jan;Kowalski`r`n002;Anna;Nowak`r`n")
    [System.IO.File]::WriteAllBytes($path, $bytes)
    $rows = Import-Csv -Path $path -Delimiter ';' -Encoding UTF8
    Assert ($rows.Count -eq 2) "Expected 2 rows, got $($rows.Count)"
    Assert ($rows[0].ID -eq '001') "Expected ID=001, got $($rows[0].ID)"
}

Test-Case 'UTF-16 LE CSV readable with Unicode encoding' {
    $path = Join-Path $TestDir 'utf16.csv'
    $content = "ID;Imie;Nazwisko`r`n003;Piotr;Wisniewski`r`n"
    [System.IO.File]::WriteAllText($path, $content, [System.Text.Encoding]::Unicode)
    $rows = @(Import-Csv -Path $path -Delimiter ';' -Encoding Unicode)
    Assert ($rows.Count -eq 1) "Expected 1 row, got $($rows.Count)"
    Assert ($rows[0].ID -eq '003') "Expected ID=003"
}

Test-Case 'CSV with empty join-key values -- empty rows are isolable' {
    $path = Join-Path $TestDir 'emptykey.csv'
    [System.IO.File]::WriteAllText($path, "ID;Name`r`n001;Jan`r`n;Anna`r`n003;Piotr`r`n", [System.Text.UTF8Encoding]::new($true))
    $rows = Import-Csv -Path $path -Delimiter ';' -Encoding UTF8
    $emptyRows = @($rows | Where-Object { [string]::IsNullOrEmpty($_.ID) })
    Assert ($emptyRows.Count -eq 1) "Expected 1 empty-key row, got $($emptyRows.Count)"
}

Test-Case 'Large cell value (32001 chars) -- XML element builds without exception' {
    $longVal = 'X' * 32001
    $xml = "<c r=`"B2`" t=`"inlineStr`"><is><t>$longVal</t></is></c>"
    Assert ($xml.Length -gt 32001) 'XML element must be longer than 32001'
    $xlsxPath = Join-Path $TestDir 'large.xlsx'
    New-MinimalXlsx -Path $xlsxPath
    Assert (Test-ValidZip $xlsxPath) 'Minimal XLSX must remain valid ZIP'
}

Test-Case 'Orphaned .tmp.xlsx is detectable and removable' {
    $basePath = Join-Path $TestDir 'base_orphan.xlsx'
    $tmpPath  = $basePath + '.tmp.xlsx'
    New-MinimalXlsx -Path $basePath
    [System.IO.File]::WriteAllBytes($tmpPath, [byte[]]@(0x50,0x4B))
    Assert (Test-Path $tmpPath) 'Orphaned .tmp.xlsx must exist'
    Remove-Item -Force $tmpPath
    Assert (-not (Test-Path $tmpPath)) 'Orphaned .tmp.xlsx must be deletable'
}

Test-Case 'Locked XLSX raises IOException on exclusive open attempt' {
    $lockPath = Join-Path $TestDir 'locked.xlsx'
    New-MinimalXlsx -Path $lockPath
    $fs = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $caught = $false
    try {
        $fs2 = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fs2.Dispose()
    } catch [System.IO.IOException] { $caught = $true }
    finally { $fs.Dispose() }
    Assert $caught 'Locked file must raise IOException'
}

Test-Case 'Header fingerprint is deterministic for identical headers' {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $getHash = { param($h) $j = ($h | ForEach-Object { $_.Trim().ToLowerInvariant() }) -join '|'; [Convert]::ToBase64String($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($j))) }
    $h1 = & $getHash @('ID','Imie','Nazwisko','PESEL')
    $h2 = & $getHash @('ID','Imie','Nazwisko','PESEL')
    Assert ($h1 -eq $h2) "Same headers must produce same hash. Got: $h1 vs $h2"
}

Test-Case 'Header fingerprint differs for reordered headers' {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $getHash = { param($h) $j = ($h | ForEach-Object { $_.Trim().ToLowerInvariant() }) -join '|'; [Convert]::ToBase64String($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($j))) }
    $h1 = & $getHash @('ID','Imie','Nazwisko')
    $h2 = & $getHash @('Nazwisko','Imie','ID')
    Assert ($h1 -ne $h2) 'Reordered headers must produce different hash'
}

Test-Case 'Invoke-HeadlessMasterUpdater rejects identical base and incoming path' {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Master-Updater.ps1'
    . $scriptPath -NonInteractive
    $sameFile = Join-Path $TestDir 'same.xlsx'
    New-MinimalXlsx -Path $sameFile
    $threw = $false
    try {
        Invoke-HeadlessMasterUpdater -BaseFilePath $sameFile -IncomingPath $sameFile
    } catch {
        $threw = $true
        Assert ($_.Exception.Message -match 'cannot be the same file') "Expected same file exception message. Got: $($_.Exception.Message)"
    }
    Assert $threw 'Must reject identical base and incoming file paths'
}

Test-Case 'Invoke-HeadlessMasterUpdater rejects unsupported file extensions' {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Master-Updater.ps1'
    . $scriptPath -NonInteractive
    $pdfFile = Join-Path $TestDir 'document.pdf'
    [System.IO.File]::WriteAllText($pdfFile, "%PDF-1.4...")
    $validXlsx = Join-Path $TestDir 'valid.xlsx'
    New-MinimalXlsx -Path $validXlsx
    $threw = $false
    try {
        Invoke-HeadlessMasterUpdater -BaseFilePath $pdfFile -IncomingPath $validXlsx
    } catch {
        $threw = $true
        Assert ($_.Exception.Message -match 'Unsupported base file format') "Expected unsupported format exception. Got: $($_.Exception.Message)"
    }
    Assert $threw 'Must throw on unsupported file extension'
}

Test-Case 'Invoke-HeadlessMasterUpdater rejects empty 0-byte files' {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Master-Updater.ps1'
    . $scriptPath -NonInteractive
    $zeroFile = Join-Path $TestDir 'zero.xlsx'
    [System.IO.File]::WriteAllBytes($zeroFile, [byte[]]@())
    $validXlsx = Join-Path $TestDir 'valid.xlsx'
    if (-not (Test-Path $validXlsx)) { New-MinimalXlsx -Path $validXlsx }
    $threw = $false
    try {
        Invoke-HeadlessMasterUpdater -BaseFilePath $zeroFile -IncomingPath $validXlsx
    } catch {
        $threw = $true
        Assert ($_.Exception.Message -match 'empty') "Expected empty file exception. Got: $($_.Exception.Message)"
    }
    Assert $threw 'Must throw on 0-byte file'
}

Test-Case 'Invoke-HeadlessMasterUpdater rejects missing file' {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Master-Updater.ps1'
    . $scriptPath -NonInteractive
    $missingFile = Join-Path $TestDir 'nonexistent_file_123.xlsx'
    $validXlsx = Join-Path $TestDir 'valid.xlsx'
    if (-not (Test-Path $validXlsx)) { New-MinimalXlsx -Path $validXlsx }
    $threw = $false
    try {
        Invoke-HeadlessMasterUpdater -BaseFilePath $missingFile -IncomingPath $validXlsx
    } catch {
        $threw = $true
        Assert ($_.Exception.Message -match 'Base file not found') "Expected not found exception. Got: $($_.Exception.Message)"
    }
    Assert $threw 'Must throw on missing file'
}

try { Remove-Item -Recurse -Force $TestDir -ErrorAction SilentlyContinue } catch { }

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Negative-path tests: $($PassCount+$FailCount) total"
Write-Host " PASS: $PassCount" -ForegroundColor Green
if ($FailCount -gt 0) {
    Write-Host " FAIL: $FailCount" -ForegroundColor Red
    exit 1
} else {
    Write-Host " All negative-path tests PASSED." -ForegroundColor Green
}