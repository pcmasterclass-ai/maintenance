<#
.SYNOPSIS
    PC Masterclass - Reschedule Maintenance Script
.DESCRIPTION
    Changes the next scheduled maintenance run date and/or time for an existing
    PCMasterclass-Maintenance scheduled task. Use this when you need to bring
    forward or delay the next scan without removing and recreating the task.

    Can also change the frequency (quarterly/bi-monthly/monthly) and run time.

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.0.1
    Date:    2026-03-16

    USAGE (paste into an elevated PowerShell prompt):
      irm https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/Reschedule-PCMasterclass.ps1 | iex
#>

# ============================================================================
# CONFIGURATION
# ============================================================================
$TaskName = "PCMasterclass-Maintenance"

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
# ADMIN CHECK
# ============================================================================
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "`n  [FAIL] This script must be run as Administrator." -ForegroundColor Red
    Write-Host "  Right-click PowerShell and select 'Run as Administrator', then try again.`n" -ForegroundColor Yellow
    Read-Host "  Press Enter to close"
    return
}

# ============================================================================
# BANNER
# ============================================================================
$banner = @"

    ====================================================
       PC Masterclass - Reschedule Maintenance v1.0.1
    ====================================================
       Change the next scheduled maintenance run
    ====================================================

"@
Write-Host $banner -ForegroundColor Cyan

# ============================================================================
# CHECK EXISTING TASK
# ============================================================================
Write-Section "CURRENT SCHEDULE"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $task) {
    Write-Fail "No scheduled task found: $TaskName"
    Write-Warn "Run the onboarding script first to create the scheduled task."
    Read-Host "`n  Press Enter to close"
    return
}

# Get current trigger details
$taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
$currentTrigger = $task.Triggers | Select-Object -First 1

# Display current settings
$nextRun = if ($taskInfo.NextRunTime) { $taskInfo.NextRunTime.ToString("dd MMM yyyy 'at' h:mmtt") } else { "Unknown" }
$lastRun = if ($taskInfo.LastRunTime -and $taskInfo.LastRunTime.Year -gt 1999) { $taskInfo.LastRunTime.ToString("dd MMM yyyy 'at' h:mmtt") } else { "Never" }

$currentInterval = 90
if ($currentTrigger.Repetition -and $currentTrigger.Repetition.Interval) {
    # Try to extract interval from trigger
    $currentInterval = 90
}
# Check DaysInterval from the Daily trigger
if ($currentTrigger -and $currentTrigger.PSObject.Properties['DaysInterval']) {
    $currentInterval = $currentTrigger.DaysInterval
}

$currentFreqLabel = switch ($currentInterval) {
    30 { "Monthly (every 30 days)" }
    60 { "Bi-monthly (every 60 days)" }
    90 { "Quarterly (every 90 days)" }
    default { "Every $currentInterval days" }
}

Write-Host ""
Write-Host "  Computer:       $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  Task status:    $($task.State)" -ForegroundColor White
Write-Host "  Frequency:      $currentFreqLabel" -ForegroundColor White
Write-Host "  Next run:       $nextRun" -ForegroundColor White
Write-Host "  Last run:       $lastRun" -ForegroundColor White

# Get the action to preserve EmailTo
$currentAction = $task.Actions | Select-Object -First 1
$currentArgs = $currentAction.Arguments

# ============================================================================
# OPTIONS
# ============================================================================
Write-Section "WHAT WOULD YOU LIKE TO CHANGE?"

Write-Host ""
Write-Host "  1. Change next run date only (keep same time and frequency)" -ForegroundColor White
Write-Host "  2. Change next run date, time, and frequency" -ForegroundColor White
Write-Host "  3. Run maintenance now (trigger an immediate run)" -ForegroundColor White
Write-Host "  4. Cancel (no changes)" -ForegroundColor White
Write-Host ""

$choice = Read-Host "  Choose an option (1-4)"

switch ($choice) {
    "1" {
        # Change date only
        Write-Host ""
        $todayStr = (Get-Date).ToString("dd/MM/yyyy")
        $dateInput = Read-Host "  New next run date dd/MM/yyyy (e.g. $todayStr)"
        if (-not $dateInput) {
            Write-Warn "No date entered - no changes made"
            Read-Host "`n  Press Enter to close"
            return
        }

        try {
            $newDate = [datetime]::ParseExact($dateInput, "dd/MM/yyyy", $null)
        } catch {
            Write-Fail "Invalid date format. Use dd/MM/yyyy (e.g. 15/04/2026)"
            Read-Host "`n  Press Enter to close"
            return
        }

        # Keep existing time from current trigger
        $existingTime = "1:00AM"
        if ($currentTrigger -and $currentTrigger.PSObject.Properties['StartBoundary']) {
            try {
                $existingDateTime = [datetime]$currentTrigger.StartBoundary
                $existingTime = $existingDateTime.ToString("h:mmtt")
            } catch {}
        }

        $newStartDateTime = $newDate.Add([datetime]::Parse($existingTime).TimeOfDay)

        # Create new trigger with same interval
        $newTrigger = New-ScheduledTaskTrigger -Daily -DaysInterval $currentInterval -At $newStartDateTime

        Set-ScheduledTask -TaskName $TaskName -Trigger $newTrigger | Out-Null

        # Verify the DaysInterval actually stuck (some Windows builds ignore it)
        $vTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        $vTrigger = $vTask.Triggers | Select-Object -First 1
        $vInterval = 1
        if ($vTrigger -and $vTrigger.PSObject.Properties['DaysInterval']) { $vInterval = $vTrigger.DaysInterval }
        if ($vInterval -ne $currentInterval) {
            Write-Warn "Windows set interval to $vInterval day(s) instead of $currentInterval - fixing via XML..."
            try {
                $xml = Export-ScheduledTask -TaskName $TaskName
                $xml = $xml -replace '<DaysInterval>\d+</DaysInterval>', "<DaysInterval>$currentInterval</DaysInterval>"
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                Register-ScheduledTask -TaskName $TaskName -Xml $xml -User "SYSTEM" -Force | Out-Null
                Write-OK "Fixed: interval corrected to every $currentInterval days"
            } catch {
                Write-Fail "Could not fix interval: $($_.Exception.Message)"
            }
        }

        Write-OK "Next run rescheduled to: $($newStartDateTime.ToString('dd MMM yyyy')) at $existingTime"
        Write-OK "Frequency unchanged: $currentFreqLabel"
    }

    "2" {
        # Change date, time, and frequency
        Write-Host ""
        $todayStr = (Get-Date).ToString("dd/MM/yyyy")
        $dateInput = Read-Host "  New next run date dd/MM/yyyy (e.g. $todayStr)"
        if (-not $dateInput) {
            Write-Warn "No date entered - no changes made"
            Read-Host "`n  Press Enter to close"
            return
        }

        try {
            $newDate = [datetime]::ParseExact($dateInput, "dd/MM/yyyy", $null)
        } catch {
            Write-Fail "Invalid date format. Use dd/MM/yyyy (e.g. 15/04/2026)"
            Read-Host "`n  Press Enter to close"
            return
        }

        Write-Host ""
        $timeInput = Read-Host "  Run time (press Enter for 1:00AM)"
        $runTime = if ($timeInput) { $timeInput } else { "1:00AM" }

        Write-Host ""
        Write-Host "  Frequency options:" -ForegroundColor White
        Write-Host "    1. Every 90 days (quarterly - recommended)" -ForegroundColor White
        Write-Host "    2. Every 60 days (bi-monthly)" -ForegroundColor White
        Write-Host "    3. Every 30 days (monthly)" -ForegroundColor White
        Write-Host ""
        $freqChoice = Read-Host "  Choose frequency (press Enter for quarterly)"

        $frequencyDays = switch ($freqChoice) {
            "2" { 60 }
            "3" { 30 }
            default { 90 }
        }
        $freqLabel = switch ($frequencyDays) {
            30 { "monthly (every 30 days)" }
            60 { "bi-monthly (every 60 days)" }
            90 { "quarterly (every 90 days)" }
        }

        $newStartDateTime = $newDate.Add([datetime]::Parse($runTime).TimeOfDay)

        $newTrigger = New-ScheduledTaskTrigger -Daily -DaysInterval $frequencyDays -At $newStartDateTime

        Set-ScheduledTask -TaskName $TaskName -Trigger $newTrigger | Out-Null

        # Verify the DaysInterval actually stuck (some Windows builds ignore it)
        $vTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        $vTrigger = $vTask.Triggers | Select-Object -First 1
        $vInterval = 1
        if ($vTrigger -and $vTrigger.PSObject.Properties['DaysInterval']) { $vInterval = $vTrigger.DaysInterval }
        if ($vInterval -ne $frequencyDays) {
            Write-Warn "Windows set interval to $vInterval day(s) instead of $frequencyDays - fixing via XML..."
            try {
                $xml = Export-ScheduledTask -TaskName $TaskName
                $xml = $xml -replace '<DaysInterval>\d+</DaysInterval>', "<DaysInterval>$frequencyDays</DaysInterval>"
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                Register-ScheduledTask -TaskName $TaskName -Xml $xml -User "SYSTEM" -Force | Out-Null
                Write-OK "Fixed: interval corrected to every $frequencyDays days"
            } catch {
                Write-Fail "Could not fix interval: $($_.Exception.Message)"
            }
        }

        Write-OK "Next run rescheduled to: $($newStartDateTime.ToString('dd MMM yyyy')) at $runTime"
        Write-OK "Frequency changed to: $freqLabel"
    }

    "3" {
        # Trigger immediate run
        Write-Step "Triggering immediate maintenance run..."
        try {
            Start-ScheduledTask -TaskName $TaskName
            Write-OK "Maintenance run started in background"
            Write-OK "Report will be emailed when complete (10-15 mins)"
        } catch {
            Write-Fail "Could not trigger run: $($_.Exception.Message)"
            Write-Warn "Try starting it manually from Task Scheduler"
        }
    }

    "4" {
        Write-Host ""
        Write-Host "  No changes made." -ForegroundColor Cyan
    }

    default {
        Write-Warn "Invalid option - no changes made"
    }
}

# ============================================================================
# SHOW UPDATED STATUS
# ============================================================================
if ($choice -in "1", "2") {
    Write-Host ""
    $updatedInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($updatedInfo.NextRunTime) {
        Write-OK "Confirmed next run: $($updatedInfo.NextRunTime.ToString('dd MMM yyyy at h:mmtt'))"
    }
}

Read-Host "`n  Press Enter to close"
