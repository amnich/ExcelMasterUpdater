#Requires -Version 5.1
<#
.SYNOPSIS
    Automated WPF GUI Acceptance Test Suite for Excel Master Updater using FlaUI.

.DESCRIPTION
    Validates end-to-end graphical user interface acceptance criteria via FlaUI (UIA3 Backend):
    1. GAC-1: STA Apartment Mode & Out-Of-Process Lifecycle (New-UiSession / Get-UiWindow).
    2. GAC-2: Main Window Visual Tree & Core Metadata Badges (winMasterUpdater, txtAppTitle).
    3. GAC-3: Dynamic Theme Switching Engine (Dark Mode <-> Light Mode via btnThemeToggle).
    4. GAC-4: Dynamic Trilingual Localization Engine (PL -> EN -> DE -> PL via cmbLanguage).
    5. GAC-5: Input Surfaces, File Choosers & Action Controls Verification.
    6. GAC-6: Resilient Out-Of-Process Teardown & Zombie Cleanup (Close-UiSession).

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7+ (pwsh -STA).
    Mandatory UTF-8 with BOM encoding.
#>

$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptDir
$exePath = Join-Path $ProjectRoot 'Master-Updater.exe'
$moduleManifest = "D:\Skrypty\WpfGuiTest\WpfGuiTest.psd1"

$PassCount = 0
$FailCount = 0

function Test-Case {
    param([string]$Name, [scriptblock]$Block)
    try {
        & $Block
        Write-Host "[GUI PASS] $Name" -ForegroundColor Green
        $script:PassCount++
    } catch {
        Write-Host "[GUI FAIL] $Name -- $($_.Exception.Message)" -ForegroundColor Red
        $script:FailCount++
    }
}

function Assert-Gui {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "GUI ACCEPTANCE VIOLATION: $Message" }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " WPF GUI ACCEPTANCE TEST SUITE (FlaUI UIA3) - EXCEL MASTER UPDATER" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# Ensure compiled executable exists
if (-not (Test-Path $exePath)) {
    throw "Compiled executable not found at $exePath. Run Build-Exe.ps1 first."
}

# Import WpfGuiTest module
Assert-Gui (Test-Path $moduleManifest) "WpfGuiTest module manifest exists at $moduleManifest"
Import-Module $moduleManifest -Force
Initialize-FlaUI
Write-Host "Loaded WpfGuiTest FlaUI automation engine." -ForegroundColor Gray

# -------------------------------------------------------------------------
# Launch Out-Of-Process UI Session
# -------------------------------------------------------------------------
Write-Host "Launching out-of-process session for Master-Updater.exe..." -ForegroundColor Gray
$session = New-UiSession -AppPath $exePath
$win = $null

try {
    # -------------------------------------------------------------------------
    # GAC-1: STA Mode & Out-Of-Process Window Attachment
    # -------------------------------------------------------------------------
    Test-Case "GAC-1: STA Mode & Out-Of-Process Window Attachment" {
        Assert-StaThread
        $script:win = Get-UiWindow -Session $session -TimeoutSeconds 12
        Assert-Gui ($null -ne $script:win) "Successfully attached to Master-Updater main window."
        Assert-Gui ($script:win.AutomationId -eq 'winMasterUpdater') "Window AutomationId asserted ('$($script:win.AutomationId)')."
    }

    # -------------------------------------------------------------------------
    # GAC-2: Main Window Visual Tree & Core Metadata Badges
    # -------------------------------------------------------------------------
    Test-Case "GAC-2: Main Window Title & Visual Branding Assertions" {
        $txtTitle = Find-UiElement -ParentElement $script:win -AutomationId "txtAppTitle"
        Assert-Gui ($txtTitle.Name -eq 'Excel Master Updater') "Application title matches ('$($txtTitle.Name)')."
        
        $txtSubtitle = Find-UiElement -ParentElement $script:win -AutomationId "txtAppSubtitle"
        Assert-Gui (-not [string]::IsNullOrWhiteSpace($txtSubtitle.Name)) "Subtitle element present and populated."
    }

    # -------------------------------------------------------------------------
    # GAC-3: Dynamic Theme Switching Engine (Dark <-> Light Mode)
    # -------------------------------------------------------------------------
    Test-Case "GAC-3: Dynamic Theme Switching Engine (Dark <-> Light Mode)" {
        $btnTheme = Find-UiElement -ParentElement $script:win -AutomationId "btnThemeToggle"
        $initialTheme = $btnTheme.Name
        Write-Host "  Initial Theme Label: $initialTheme" -ForegroundColor Gray
        
        # Click once to toggle theme
        Invoke-UiElement -Element $btnTheme -Action Click
        Start-Sleep -Milliseconds 500
        $toggledTheme = $btnTheme.Name
        Write-Host "  Toggled Theme Label: $toggledTheme" -ForegroundColor Gray
        Assert-Gui ($initialTheme -ne $toggledTheme) "Theme label dynamically changed upon click."
        
        # Click again to restore original theme
        Invoke-UiElement -Element $btnTheme -Action Click
        Start-Sleep -Milliseconds 500
        $restoredTheme = $btnTheme.Name
        Write-Host "  Restored Theme Label: $restoredTheme" -ForegroundColor Gray
        Assert-Gui ($restoredTheme -eq $initialTheme) "Original theme restored successfully."
    }

    # -------------------------------------------------------------------------
    # GAC-4: Dynamic Trilingual Localization Engine (PL -> EN -> DE -> PL)
    # -------------------------------------------------------------------------
    Test-Case "GAC-4: Dynamic Trilingual Localization Engine without App Restart" {
        $cmbLang = Find-UiElement -ParentElement $script:win -AutomationId "cmbLanguage"
        $lblLang = Find-UiElement -ParentElement $script:win -AutomationId "lblLanguage"
        $btnSettings = Find-UiElement -ParentElement $script:win -AutomationId "btnSettings"
        
        # 1. Switch to English
        $itemEn = $cmbLang.FindAllChildren() | Where-Object { $_.Name -eq "English" }
        Assert-Gui ($null -ne $itemEn) "English selection item resolved in language dropdown."
        $itemEn.Patterns.SelectionItem.Pattern.Select()
        Start-Sleep -Milliseconds 500
        Assert-Gui ($lblLang.Name -eq 'Language:') "Language label updated to English ('$($lblLang.Name)')."
        Assert-Gui ($btnSettings.Name -like '*Settings*') "Settings button updated to English ('$($btnSettings.Name)')."
        
        # 2. Switch to German
        $itemDe = $cmbLang.FindAllChildren() | Where-Object { $_.Name -eq "Deutsch" }
        Assert-Gui ($null -ne $itemDe) "Deutsch selection item resolved in language dropdown."
        $itemDe.Patterns.SelectionItem.Pattern.Select()
        Start-Sleep -Milliseconds 500
        Assert-Gui ($lblLang.Name -eq 'Sprache:') "Language label updated to German ('$($lblLang.Name)')."
        Assert-Gui ($btnSettings.Name -like '*Einstellungen*') "Settings button updated to German ('$($btnSettings.Name)')."
        
        # 3. Restore to Polish
        $itemPl = $cmbLang.FindAllChildren() | Where-Object { $_.Name -eq "Polski" }
        Assert-Gui ($null -ne $itemPl) "Polski selection item resolved in language dropdown."
        $itemPl.Patterns.SelectionItem.Pattern.Select()
        Start-Sleep -Milliseconds 500
        Assert-Gui ($lblLang.Name -eq 'Język:') "Language label restored to Polish ('$($lblLang.Name)')."
        Assert-Gui ($btnSettings.Name -like '*Ustawienia*') "Settings button restored to Polish ('$($btnSettings.Name)')."
    }

    # -------------------------------------------------------------------------
    # GAC-5: Input Surfaces, File Choosers & Action Controls Verification
    # -------------------------------------------------------------------------
    Test-Case "GAC-5: Input Surfaces, File Choosers & Action Controls Verification" {
        $txtBasePath = Find-UiElement -ParentElement $script:win -AutomationId "txtBasePath"
        Assert-Gui ($null -ne $txtBasePath) "Base file path display textbox present."
        Assert-Gui ($txtBasePath.Patterns.Value.Pattern.IsReadOnly.Value -eq $true) "Base file path textbox is guarded read-only."
        
        $btnBrowseBase = Find-UiElement -ParentElement $script:win -AutomationId "btnBrowseBase"
        Assert-Gui ($null -ne $btnBrowseBase) "Browse base file button present."
        
        $btnOpenBackups = Find-UiElement -ParentElement $script:win -AutomationId "btnOpenBackups"
        Assert-Gui ($null -ne $btnOpenBackups) "Open backups button present."
        
        $btnOpenLogs = Find-UiElement -ParentElement $script:win -AutomationId "btnOpenLogs"
        Assert-Gui ($null -ne $btnOpenLogs) "Open logs button present."
    }

} finally {
    # -------------------------------------------------------------------------
    # GAC-6: Safe Out-Of-Process Termination & Zombie Cleanup
    # -------------------------------------------------------------------------
    Test-Case "GAC-6: Safe Out-Of-Process Termination & Zombie Cleanup" {
        $procId = $session.Application.ProcessId
        Close-UiSession -Session $session
        Start-Sleep -Milliseconds 600
        $running = Get-Process -Id $procId -ErrorAction SilentlyContinue
        Assert-Gui ($null -eq $running) "Process PID $procId completely terminated without lingering zombies."
    }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " WPF GUI ACCEPTANCE TEST SUMMARY" -ForegroundColor Cyan
Write-Host " Total Criteria: $($PassCount + $FailCount) | Passed: $PassCount | Failed: $FailCount" -ForegroundColor $(if ($FailCount -eq 0) { 'Green' } else { 'Red' })
Write-Host "================================================================================" -ForegroundColor Cyan

if ($FailCount -gt 0) {
    throw "$FailCount GUI Acceptance Criteria FAILED."
}
