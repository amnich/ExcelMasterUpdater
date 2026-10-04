#Requires -Version 5.1
<#
.SYNOPSIS
    Folder-watch daemon and auto-processor for Excel Master Updater (SWOT P4 / O1).

.DESCRIPTION
    Monitors an incoming drop folder for incoming Excel (.xlsx) or CSV (.csv) files.
    When a new file arrives and is completely written (lock released), it triggers
    Master-Updater in headless batch mode, executes automated comparison and optional
    auto-accept write-back, optionally generates an email-ready HTML diff report,
    and moves the processed source file to an Archive (or Failed) folder.

.PARAMETER FolderPath
    Directory path to monitor for incoming files. Defaults to '.\IncomingDrop'.

.PARAMETER BaseFilePath
    Path to the master base Excel (.xlsx) or CSV (.csv) file to update.

.PARAMETER BaseSheet
    Name of worksheet in the base file. Defaults to first worksheet.

.PARAMETER ArchiveFolder
    Folder to move successfully processed files to. Defaults to '<FolderPath>\Archive'.

.PARAMETER FailedFolder
    Folder to move files that encountered errors to. Defaults to '<FolderPath>\Failed'.

.PARAMETER ReportsFolder
    Folder to save generated HTML/XLSX diff reports to. Defaults to '<FolderPath>\Reports'.

.PARAMETER AutoAccept
    Batch acceptance policy. Options: 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew', 'None'.
    Defaults to 'AllNonAmbiguous'.

.PARAMETER ExportHtmlReports
    Generates a standalone responsive HTML diff report for every processed file.

.PARAMETER PollIntervalSeconds
    Seconds between directory polling scans. Defaults to 3 seconds.

.PARAMETER Once
    Processes any files currently in FolderPath once and immediately exits.
    Ideal for scheduled tasks, batch runners, and automated tests.

.PARAMETER MaxIterations
    Optional maximum number of polling iterations before stopping.

.EXAMPLE
    pwsh -File .\Watch-IncomingFolder.ps1 -BaseFilePath "C:\Data\Master.xlsx" -FolderPath "C:\DropZone"
    Monitors C:\DropZone indefinitely, auto-updating Master.xlsx and moving processed files to Archive.

.EXAMPLE
    pwsh -File .\Watch-IncomingFolder.ps1 -BaseFilePath ".\base.xlsx" -FolderPath ".\Drop" -Once -ExportHtmlReports
    Processes all pending files in .\Drop once, exports HTML diff reports to .\Drop\Reports, and exits.

.OUTPUTS
    System.Management.Automation.PSCustomObject summary per processed file.

.NOTES
    Compatible with Windows PowerShell 5.1 (.NET Framework) and PowerShell 7+ (.NET Core).
    Zero external dependencies.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$FolderPath,

    [Parameter(Mandatory = $false)]
    [string]$BaseFilePath,

    [Parameter(Mandatory = $false)]
    [string]$BaseSheet,

    [Parameter(Mandatory = $false)]
    [string]$ArchiveFolder,

    [Parameter(Mandatory = $false)]
    [string]$FailedFolder,

    [Parameter(Mandatory = $false)]
    [string]$ReportsFolder,

    [Parameter(Mandatory = $false)]
    [ValidateSet('AllNonAmbiguous', 'OnlyChanged', 'OnlyNew', 'None')]
    [string]$AutoAccept = 'AllNonAmbiguous',

    [Parameter(Mandatory = $false)]
    [switch]$ExportHtmlReports,

    [Parameter(Mandatory = $false)]
    [int]$PollIntervalSeconds = 3,

    [Parameter(Mandatory = $false)]
    [switch]$Once,

    [Parameter(Mandatory = $false)]
    [int]$MaxIterations = 0
)

$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# Resolve default paths
if (-not $FolderPath) {
    $FolderPath = Join-Path $ScriptDir 'IncomingDrop'
}
if (-not [System.IO.Path]::IsPathRooted($FolderPath)) {
    $FolderPath = [System.IO.Path]::GetFullPath((Join-Path $ScriptDir $FolderPath))
}

if (-not $ArchiveFolder) {
    $ArchiveFolder = Join-Path $FolderPath 'Archive'
}
if (-not $FailedFolder) {
    $FailedFolder = Join-Path $FolderPath 'Failed'
}
if (-not $ReportsFolder) {
    $ReportsFolder = Join-Path $FolderPath 'Reports'
}

# Ensure folders exist
foreach ($dir in @($FolderPath, $ArchiveFolder, $FailedFolder, $ReportsFolder)) {
    if (-not (Test-Path $dir)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
    }
}

# Dot-source Master-Updater engine (preserve local parameters across dot-sourcing)
$masterUpdaterScript = Join-Path $ScriptDir 'Master-Updater.ps1'
if (-not (Test-Path $masterUpdaterScript)) {
    throw "Master-Updater.ps1 not found at: $masterUpdaterScript"
}
$savedParams = @{
    BaseFilePath        = $BaseFilePath
    BaseSheet           = $BaseSheet
    FolderPath          = $FolderPath
    ArchiveFolder       = $ArchiveFolder
    FailedFolder        = $FailedFolder
    ReportsFolder       = $ReportsFolder
    AutoAccept          = $AutoAccept
    Once                = $Once
    PollIntervalSeconds = $PollIntervalSeconds
    MaxIterations       = $MaxIterations
    ExportHtmlReports   = $ExportHtmlReports
}
. $masterUpdaterScript -NonInteractive
$BaseFilePath        = $savedParams.BaseFilePath
$BaseSheet           = $savedParams.BaseSheet
$FolderPath          = $savedParams.FolderPath
$ArchiveFolder       = $savedParams.ArchiveFolder
$FailedFolder        = $savedParams.FailedFolder
$ReportsFolder       = $savedParams.ReportsFolder
$AutoAccept          = $savedParams.AutoAccept
$Once                = $savedParams.Once
$PollIntervalSeconds = $savedParams.PollIntervalSeconds
$MaxIterations       = $savedParams.MaxIterations
$ExportHtmlReports   = $savedParams.ExportHtmlReports

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " EXCEL MASTER UPDATER - FOLDER WATCH DAEMON (P4 / O1)" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "Watch Folder:    $FolderPath"
Write-Host "Base File:       $(if ($BaseFilePath) { $BaseFilePath } else { '[Config / Auto]' })"
Write-Host "Archive Folder:  $ArchiveFolder"
Write-Host "Failed Folder:   $FailedFolder"
Write-Host "Reports Folder:  $ReportsFolder"
Write-Host "AutoAccept Mode: $AutoAccept"
Write-Host "Poll Interval:   $PollIntervalSeconds s | Once Mode: $Once"
Write-Host "================================================================================" -ForegroundColor Cyan

<#
.SYNOPSIS
    Tests whether a file is completely written and unlocked by attempting exclusive read access.
#>
function Test-FileReady {
    param([string]$FilePath)
    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Dispose()
        return $true
    } catch {
        return $false
    }
}

<#
.SYNOPSIS
    Processes a single incoming Excel or CSV file.
#>
Set-Alias -Name 'Invoke-IncomingFileProcessing' -Value 'Process-IncomingFile' -ErrorAction SilentlyContinue
function Process-IncomingFile {
    param(
        [string]$IncomingFilePath,
        [string]$TargetBaseFilePath,
        [string]$TargetBaseSheet
    )

    $fileName = [System.IO.Path]::GetFileName($IncomingFilePath)
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Processing incoming file: $fileName" -ForegroundColor Green

    # Wait for write lock to clear if still copying
    $retries = 5
    while (-not (Test-FileReady $IncomingFilePath) -and $retries -gt 0) {
        Write-Host "  Waiting for file lock release on $fileName..." -ForegroundColor Gray
        Start-Sleep -Milliseconds 800
        $retries--
    }

    if (-not (Test-FileReady $IncomingFilePath)) {
        Write-Warning "File $fileName is locked by another process. Skipping for this cycle."
        return $null
    }

    # Resolve Base file if not passed
    $effectiveBase = $TargetBaseFilePath
    if (-not $effectiveBase) {
        $cfg = Get-AppConfig
        $effectiveBase = $cfg.BaseFilePath
    }
    if (-not $effectiveBase -or -not (Test-Path $effectiveBase)) {
        Write-Warning "No valid BaseFilePath found for processing $fileName."
        return $null
    }

    # Destination report path if requested
    $reportPath = $null
    if ($ExportHtmlReports) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $reportPath = Join-Path $ReportsFolder "${baseName}_diff_${ts}.html"
    }

    try {
        $result = Invoke-HeadlessMasterUpdater -BaseFilePath $effectiveBase -IncomingPath $IncomingFilePath -BaseSheet $TargetBaseSheet -AutoAccept $AutoAccept -ExportReportPath $reportPath

        # Archive processed file
        $archiveTarget = Join-Path $ArchiveFolder "${ts}_${fileName}"
        Move-Item -Path $IncomingFilePath -Destination $archiveTarget -Force
        Write-Host "  Moved $fileName -> Archive" -ForegroundColor Gray

        return [PSCustomObject]@{
            FileName     = $fileName
            Success      = $true
            Total        = $result.TotalIncoming
            New          = $result.NewRows
            Changed      = $result.ChangedRows
            Ambiguous    = $result.AmbiguousRows
            Accepted     = $result.AcceptedRows
            UpdatedCells = $result.UpdatedCells
            AddedRows    = $result.AddedRows
            Report       = $reportPath
            SummaryJson  = $result.SummaryJsonPath
            Timestamp    = $ts
        }
    } catch {
        Write-Host "  ERROR processing $fileName`: $($_.Exception.Message)" -ForegroundColor Red
        # Move to Failed folder
        $failedTarget = Join-Path $FailedFolder "${ts}_${fileName}"
        try {
            Move-Item -Path $IncomingFilePath -Destination $failedTarget -Force
            Write-Host "  Moved $fileName -> Failed" -ForegroundColor DarkYellow
        } catch {
            Write-Warning "Failed to move '$fileName' to '$failedTarget': $($_.Exception.Message)"
        }

        return [PSCustomObject]@{
            FileName     = $fileName
            Success      = $false
            ErrorMessage = $_.Exception.Message
            Timestamp    = $ts
        }
    }
}

$iteration = 0
$processedResults = [System.Collections.Generic.List[object]]::new()

try {
    while ($true) {
        $iteration++
        if ($MaxIterations -gt 0 -and $iteration -gt $MaxIterations) {
            Write-Host "Reached max iterations ($MaxIterations). Exiting watcher." -ForegroundColor Cyan
            break
        }

        # Scan for candidate files (exclude subfolders)
        $candidates = @(Get-ChildItem -Path $FolderPath -File | Where-Object {
            $_.Extension -match '^\.(xlsx|csv)$' -and
            -not $_.Name.StartsWith('~') -and
            -not $_.Name.EndsWith('.tmp.xlsx')
        })

        if ($candidates.Count -gt 0) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Found $($candidates.Count) file(s) in drop folder." -ForegroundColor Yellow
            foreach ($file in $candidates) {
                $res = Process-IncomingFile -IncomingFilePath $file.FullName -TargetBaseFilePath $BaseFilePath -TargetBaseSheet $BaseSheet
                if ($res) {
                    $processedResults.Add($res)
                }
            }
        }

        if ($Once) {
            Write-Host "Completed one-shot run (-Once). Exiting." -ForegroundColor Cyan
            break
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }
} catch [System.Management.Automation.PipelineStoppedException] {
    Write-Host "`nFolder watcher stopped by user." -ForegroundColor Yellow
}

return $processedResults
