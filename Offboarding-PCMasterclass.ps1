<#
.SYNOPSIS
    PC Masterclass - Client Offboarding Script
.DESCRIPTION
    Removes the PC Masterclass maintenance environment from a client machine.
    Run this via TeamViewer (or in person) when a client no longer wishes to
    continue with the maintenance service.

    Removes:
      - The scheduled maintenance task
      - Stored email credentials (DPAPI + AES)
      - The maintenance script
      - Optionally: the entire C:\Teamviewer folder (with reports)

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.0.1
    Date:    2026-03-12

    USAGE (paste into an elevated PowerShell prompt):
      powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\Offboarding-PCMasterclass.ps1"
#>

# ============================================================================
# CONFIGURATION
# ============================================================================
$BaseDir    = "C:\Teamviewer"
$ConfigDir  = Join-Path $BaseDir "Config"
$TaskName   = "PCMasterclass-Maintenance"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
function Write-Step {
    param([string]$Message)
    Write-Host "  [*] $Message" -ForegroundColor White
}

function Write-OK {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "  [!!] $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  [FAIL] $Message" -ForegroundColor Red
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "  --- $Title ---" -ForegroundColor Cyan
}

# ============================================================================
# BANNER
# ============================================================================
# Check for admin rights (replaces #Requires -RunAsAdministrator which breaks irm | iex)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "`n  [FAIL] This script must be run as Administrator." -ForegroundColor Red
    Write-Host "  Right-click PowerShell and select 'Run as Administrator', then try again.`n" -ForegroundColor Yellow
    Read-Host "  Press Enter to close"
    return
}

$banner = @"

    ====================================================
       PC Masterclass - Client Offboarding Script v1.0.1
    ====================================================
       Removing maintenance environment from this machine
    ====================================================

"@
Write-Host $banner -ForegroundColor Yellow

# ============================================================================
# CONFIRMATION
# ============================================================================
Write-Host "  This will remove the PC Masterclass maintenance setup from" -ForegroundColor White
Write-Host "  this computer ($env:COMPUTERNAME)." -ForegroundColor White
Write-Host ""
Write-Host "  The following will be removed:" -ForegroundColor White
Write-Host "    - Scheduled maintenance task" -ForegroundColor White
Write-Host "    - Stored email credentials" -ForegroundColor White
Write-Host "    - Maintenance script" -ForegroundColor White
Write-Host ""

$confirm = Read-Host "  Are you sure you want to proceed? (Y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host ""
    Write-Host "  Offboarding cancelled." -ForegroundColor Cyan
    Read-Host "`n  Press Enter to close"
    return
}

# ============================================================================
# STEP 1: REMOVE SCHEDULED TASK
# ============================================================================
Write-Section "REMOVING SCHEDULED TASK"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-OK "Scheduled task removed: $TaskName"
    } catch {
        Write-Fail "Failed to remove scheduled task: $($_.Exception.Message)"
    }
} else {
    Write-OK "No scheduled task found (already removed or never created)"
}

# ============================================================================
# STEP 2: REMOVE CREDENTIALS
# ============================================================================
Write-Section "REMOVING STORED CREDENTIALS"

$credFiles = @(
    (Join-Path $ConfigDir "smtp-cred.xml"),
    (Join-Path $ConfigDir "smtp.key"),
    (Join-Path $ConfigDir "smtp-config.xml")
)

$removed = 0
foreach ($file in $credFiles) {
    if (Test-Path $file) {
        try {
            Remove-Item $file -Force -ErrorAction Stop
            Write-OK "Removed: $file"
            $removed++
        } catch {
            Write-Fail "Failed to remove: $file - $($_.Exception.Message)"
        }
    }
}
if ($removed -eq 0) {
    Write-OK "No credential files found"
}

# ============================================================================
# STEP 3: REMOVE MAINTENANCE SCRIPT
# ============================================================================
Write-Section "REMOVING MAINTENANCE SCRIPT"

$scriptPath = Join-Path $BaseDir "PCMasterclass-Maintenance.ps1"
if (Test-Path $scriptPath) {
    try {
        Remove-Item $scriptPath -Force -ErrorAction Stop
        Write-OK "Removed: $scriptPath"
    } catch {
        Write-Fail "Failed to remove: $scriptPath - $($_.Exception.Message)"
    }
} else {
    Write-OK "Maintenance script already removed"
}

# ============================================================================
# STEP 4: OPTIONAL - REMOVE ENTIRE FOLDER
# ============================================================================
Write-Section "CLEANUP"

if (Test-Path $BaseDir) {
    # Check if there are any reports
    $reportsDir = Join-Path $BaseDir "Reports"
    $reportCount = 0
    if (Test-Path $reportsDir) {
        $reportCount = (Get-ChildItem $reportsDir -File -ErrorAction SilentlyContinue | Measure-Object).Count
    }

    if ($reportCount -gt 0) {
        Write-Host ""
        Write-Host "  There are $reportCount report(s) in $reportsDir" -ForegroundColor Yellow
        Write-Host "  These contain maintenance history for this machine." -ForegroundColor White
        Write-Host ""
        $deleteAll = Read-Host "  Delete the entire $BaseDir folder including reports? (Y/N)"
    } else {
        $deleteAll = Read-Host "  Delete the $BaseDir folder? (Y/N)"
    }

    if ($deleteAll -match '^[Yy]') {
        try {
            Remove-Item $BaseDir -Recurse -Force -ErrorAction Stop
            Write-OK "Removed: $BaseDir (and all contents)"
        } catch {
            Write-Fail "Failed to remove folder: $($_.Exception.Message)"
            Write-Warn "You may need to manually delete $BaseDir"
        }
    } else {
        Write-Warn "Folder kept: $BaseDir"
        # Still clean up empty subdirectories
        foreach ($sub in @("Config", "Tools")) {
            $subPath = Join-Path $BaseDir $sub
            if ((Test-Path $subPath) -and (Get-ChildItem $subPath -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
                Remove-Item $subPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
} else {
    Write-OK "Folder already removed: $BaseDir"
}

# ============================================================================
# SUMMARY
# ============================================================================
$summary = @"

    ====================================================
       OFFBOARDING COMPLETE
    ====================================================

       Computer: $env:COMPUTERNAME

       The PC Masterclass maintenance service has been
       removed from this machine.

       Remember to update the Rollout Tracker status
       to reflect this client's removal.

    ====================================================
"@
Write-Host $summary -ForegroundColor Green

Read-Host "`n  Press Enter to close"
