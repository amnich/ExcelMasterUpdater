#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles Master-Updater.ps1 into a standalone executable (.EXE) using the PS2EXE module.

.DESCRIPTION
    Produces a windowed application without console (-noConsole) in Single-Threaded Apartment (-STA) mode.
    Embeds the trilingual language catalog (language.json) via -embedFiles and sets application metadata.
    Since Master-Updater utilizes built-in .NET standard libraries (System.IO.Compression & System.Xml),
    the resulting executable is completely standalone with zero external dependencies.

.PARAMETER OutputFile
    Path to the output executable (.EXE). Defaults to Master-Updater.exe next to this script.

.PARAMETER IconFile
    Path to a .ico icon file. Defaults to D:\Skrypty\Mnich_Adam_Skrypty\!Helper\Logo_AM6.ico.

.EXAMPLE
    powershell.exe -File .\Build-Exe.ps1
    Compiles Master-Updater.ps1 into .\Master-Updater.exe with default icon and embedded language.json.

.EXAMPLE
    powershell.exe -File .\Build-Exe.ps1 -OutputFile "C:\Tools\Updater.exe"
    Compiles the application to a custom target destination path.

.OUTPUTS
    System.IO.FileInfo. The compiled standalone binary executable.

.NOTES
    Requires PS2EXE module installed (Install-Module ps2exe -Scope CurrentUser).
    Dual-engine build script verified under Windows PowerShell 5.1 and PowerShell 7.

.LINK
    https://github.com/MScholtes/PS2EXE
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$IconFile = "D:\Skrypty\Mnich_Adam_Skrypty\!Helper\Logo_AM6.ico"
)

$ErrorActionPreference = 'Stop'
$ScriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } elseif ($MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
$InputScript = Join-Path $ScriptDir 'Master-Updater.ps1'

if (-not $OutputFile) {
    $OutputFile = Join-Path $ScriptDir 'Master-Updater.exe'
}

$AppTitle = 'Excel Master Updater - Master File Synchronizer'
$AppDesc  = 'High-performance Excel master updater with row-by-row review, InPlace OpenXML patching, and zero external dependencies'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " PS2EXE COMPILATION: Master-Updater" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Close any running instance of the output executable to avoid file lock
$targetProcName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
$runningProc = Get-Process -Name $targetProcName -ErrorAction SilentlyContinue
if ($runningProc) {
    Write-Host "Stopping running instance of $targetProcName (PID: $($runningProc.Id))..." -ForegroundColor Yellow
    $runningProc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

$ps2exeParams = @{
    InputFile     = $InputScript
    OutputFile    = $OutputFile
    Title         = $AppTitle
    Description   = $AppDesc
    Company       = 'Adam Mnich'
    Product       = 'Excel Master Updater'
    Copyright     = "Copyright (c) 2026 Adam Mnich"
    Version       = '1.0.0.0'
    noConsole     = $true
    STA           = $true
    x64           = $true
}

$embedFiles = @{}
$langFile = Join-Path $ScriptDir 'language.json'
if (Test-Path $langFile) {
    $embedFiles['.\language.json'] = $langFile
    Write-Host "Embedding language catalog: $langFile" -ForegroundColor Gray
}

$docsDir = Join-Path $ScriptDir 'Docs'
if (Test-Path $docsDir) {
    Get-ChildItem -LiteralPath $docsDir -File | ForEach-Object {
        $embedFiles[".\Docs\$($_.Name)"] = $_.FullName
    }
    Write-Host "Embedding documentation files: $docsDir" -ForegroundColor Gray
}

if ($embedFiles.Count -gt 0) {
    $ps2exeParams['embedFiles'] = $embedFiles
}

if ($IconFile -and (Test-Path $IconFile)) {
    $ps2exeParams['IconFile'] = $IconFile
    Write-Host "Using icon: $IconFile" -ForegroundColor Gray
}

Write-Host "Compiling $InputScript -> $OutputFile..." -ForegroundColor Green
Invoke-ps2exe @ps2exeParams

if (Test-Path $OutputFile) {
    $sizeKb = [Math]::Round((Get-Item $OutputFile).Length / 1KB, 1)
    Write-Host "Build SUCCESS: $OutputFile ($sizeKb KB)" -ForegroundColor Green
} else {
    throw "Build failed: output file not found."
}
