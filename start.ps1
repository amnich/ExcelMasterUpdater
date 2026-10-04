<#
.SYNOPSIS
    Excel Master Updater Launcher.

.DESCRIPTION
    Entry point launcher script for Excel Master Updater.
    Ensures STA apartment mode for WPF, supports the -UseExe switch to delegate
    to the compiled standalone binary, and forwards file and configuration parameters.

.PARAMETER BaseFilePath
    Absolute or relative path to the master base Excel file (.xlsx) or CSV file (.csv).
    Pre-populates the base file selector upon application launch.

.PARAMETER IncomingPath
    Absolute or relative path to the incoming updates Excel file (.xlsx) or CSV file (.csv).
    Pre-populates the incoming file selector upon application launch.

.PARAMETER BaseSheet
    Name of the worksheet in the base file to target for comparison and write-back.
    Defaults to the first available data worksheet if omitted.

.PARAMETER IncomingSheet
    Name of the worksheet in the incoming file to target for comparison.
    Defaults to the first available data worksheet if omitted.

.PARAMETER ConfigPath
    Custom path to the application configuration file (config.json).
    Defaults to '%APPDATA%\MasterUpdater\config.json' if omitted.

.PARAMETER NonInteractive
    Initializes the engine, dependencies, and WPF controls in headless mode without
    calling ShowDialog(). Used for automated regression testing and pipeline verification.

.PARAMETER UseExe
    When specified, launches the compiled standalone binary (Master-Updater.exe) instead
    of running the raw PowerShell script.

.EXAMPLE
    pwsh -File .\start.ps1
    Launches the Excel Master Updater graphical interface using PowerShell 7.

.EXAMPLE
    powershell.exe -File .\start.ps1 -BaseFilePath "C:\Data\Master.xlsx" -IncomingPath "C:\Data\Sept.xlsx"
    Launches the application with pre-loaded base and incoming files in Windows PowerShell 5.1.

.EXAMPLE
    pwsh -File .\start.ps1 -UseExe
    Launches the compiled standalone binary (Master-Updater.exe).

.OUTPUTS
    System.Void. Displays interactive WPF window or exits with 0 on headless verification.

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7.x.
    Requires Single-Threaded Apartment (STA) state for WPF UI rendering.

.LINK
    https://github.com/
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BaseFilePath,

    [Parameter(Mandatory = $false)]
    [string]$IncomingPath,

    [Parameter(Mandatory = $false)]
    [string]$BaseSheet,

    [Parameter(Mandatory = $false)]
    [string]$IncomingSheet,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive,

    [Parameter(Mandatory = $false)]
    [switch]$Headless,

    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew')]
    [string]$AutoAccept = 'None',

    [Parameter(Mandatory = $false)]
    [string]$ExportReportPath,

    [Parameter(Mandatory = $false)]
    [switch]$Quiet,

    [Parameter(Mandatory = $false)]
    [switch]$UseExe
)

$exeFile = Join-Path $PSScriptRoot "Master-Updater.exe"
if ($UseExe -and (Test-Path $exeFile)) {
    $exeArgs = [System.Collections.Generic.List[string]]::new()
    if ($BaseFilePath)     { $exeArgs.Add("-BaseFilePath `"$BaseFilePath`"") }
    if ($IncomingPath)     { $exeArgs.Add("-IncomingPath `"$IncomingPath`"") }
    if ($BaseSheet)        { $exeArgs.Add("-BaseSheet `"$BaseSheet`"") }
    if ($IncomingSheet)    { $exeArgs.Add("-IncomingSheet `"$IncomingSheet`"") }
    if ($ConfigPath)       { $exeArgs.Add("-ConfigPath `"$ConfigPath`"") }
    if ($NonInteractive)   { $exeArgs.Add("-NonInteractive") }
    if ($Headless)         { $exeArgs.Add("-Headless") }
    if ($AutoAccept)       { $exeArgs.Add("-AutoAccept $AutoAccept") }
    if ($ExportReportPath) { $exeArgs.Add("-ExportReportPath `"$ExportReportPath`"") }
    if ($Quiet)            { $exeArgs.Add("-Quiet") }

    Start-Process -FilePath $exeFile -ArgumentList ($exeArgs -join ' ')
    return
}

# STA Apartment check for Windows PowerShell 5.1
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    if ($PSVersionTable.PSVersion.Major -le 5) {
        Write-Host "Re-launching in STA apartment mode for WPF..." -ForegroundColor Yellow
        $argList = @("-NoProfile", "-STA", "-File", $MyInvocation.MyCommand.Path)
        if ($BaseFilePath)     { $argList += @("-BaseFilePath", $BaseFilePath) }
        if ($IncomingPath)     { $argList += @("-IncomingPath", $IncomingPath) }
        if ($BaseSheet)        { $argList += @("-BaseSheet", $BaseSheet) }
        if ($IncomingSheet)    { $argList += @("-IncomingSheet", $IncomingSheet) }
        if ($ConfigPath)       { $argList += @("-ConfigPath", $ConfigPath) }
        if ($NonInteractive)   { $argList += "-NonInteractive" }
        if ($Headless)         { $argList += "-Headless" }
        if ($AutoAccept)       { $argList += @("-AutoAccept", $AutoAccept) }
        if ($ExportReportPath) { $argList += @("-ExportReportPath", $ExportReportPath) }
        if ($Quiet)            { $argList += "-Quiet" }
        & powershell.exe @argList
        return
    }
}

$scriptFile = Join-Path $PSScriptRoot "Master-Updater.ps1"
if (-not (Test-Path $scriptFile)) {
    throw "Master-Updater.ps1 not found in: $PSScriptRoot"
}

$params = @{}
if ($BaseFilePath)     { $params['BaseFilePath']     = $BaseFilePath }
if ($IncomingPath)     { $params['IncomingPath']     = $IncomingPath }
if ($BaseSheet)        { $params['BaseSheet']        = $BaseSheet }
if ($IncomingSheet)    { $params['IncomingSheet']    = $IncomingSheet }
if ($ConfigPath)       { $params['ConfigPath']       = $ConfigPath }
if ($NonInteractive)   { $params['NonInteractive']   = $true }
if ($Headless)         { $params['Headless']         = $true }
if ($AutoAccept)       { $params['AutoAccept']       = $AutoAccept }
if ($ExportReportPath) { $params['ExportReportPath'] = $ExportReportPath }
if ($Quiet)            { $params['Quiet']            = $true }

try {
    & $scriptFile @params
} catch {
    $crashLog = Join-Path $PSScriptRoot 'crash.log'
    $info = "LAUNCHER CRASH:`n" +
            "Message: $($_.Exception.Message)`n" +
            "ScriptStackTrace:`n$($_.ScriptStackTrace)`n" +
            "Position:`n$($_.InvocationInfo.PositionMessage)`n" +
            "Inner: $($_.Exception.InnerException)`n"
    [System.IO.File]::WriteAllText($crashLog, $info, [System.Text.UTF8Encoding]::new($true))
    Write-Host $info -ForegroundColor Red
    throw
}
