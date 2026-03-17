<#
.SYNOPSIS
    PC Masterclass - Enable System Restore Points
.DESCRIPTION
    Enables Windows System Protection (restore points) on the C: drive after
    checking available disk space. If free space is below the safety threshold,
    the script logs a warning and exits without enabling — flagging the machine
    as a candidate for a storage conversation with the client.

    Safe to run on machines where restore points are already enabled — it will
    detect the existing configuration and skip gracefully.

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.1.0
    Date:    2026-03-18

    USAGE (Tactical RMM - run as SYSTEM):
      powershell -ExecutionPolicy Bypass -File "C:\Teamviewer\Enable-RestorePoints-PCMasterclass.ps1"

    THRESHOLDS:
      - Minimum free space to enable:  30 GB  (won't enable if below this)
      - Restore point allocation:       5%    (of total C: drive capacity)
      - Maximum allocation cap:        10%    (for drives over 500 GB)
#>

# ── Configuration ──────────────────────────────────────────────────────────────
$MinFreeSpaceGB       = 30          # Don't enable if free space is below this
$AllocationPercent    = 5           # Percentage of drive to allocate for restore points
$MaxAllocationPercent = 10          # Cap for very large drives
$Drive                = "C:\"
$LogPath              = "C:\Teamviewer\Reports"
$LogFile              = Join-Path $LogPath "RestorePointSetup.log"

# ── Helpers ────────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    Write-Host $entry
    if (Test-Path $LogPath) {
        $entry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}

# ── Admin Check ────────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "Script must run as Administrator / SYSTEM. Exiting." "ERROR"
    exit 1
}

# ── Check Current Status ──────────────────────────────────────────────────────
Write-Log "Checking System Protection status on C: drive..."

try {
    $restoreStatus = (Get-ComputerRestorePoint -ErrorAction SilentlyContinue) -ne $null
} catch {
    $restoreStatus = $false
}

# More reliable check via WMI
$shadowStorage = $null
try {
    $shadowStorage = Get-CimInstance -ClassName Win32_ShadowStorage -ErrorAction SilentlyContinue |
                     Where-Object { $_.Volume.DeviceID -like "*C:*" -or $_.DiffVolume.DeviceID -like "*C:*" }
} catch {}

$protectionEnabled = $false
try {
    $vssOutput = vssadmin list shadowstorage 2>&1 | Out-String
    if ($vssOutput -match "For volume: \(C:\)") {
        $protectionEnabled = $true
    }
} catch {}

# Also check via registry
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
$regDisabled = (Get-ItemProperty -Path $regPath -Name "RPSessionInterval" -ErrorAction SilentlyContinue).RPSessionInterval
$regDisabledExplicit = (Get-ItemProperty -Path $regPath -Name "DisableSR" -ErrorAction SilentlyContinue).DisableSR

if ($protectionEnabled -and $regDisabledExplicit -ne 1) {
    Write-Log "System Protection is ALREADY ENABLED on C: drive. No changes needed." "INFO"
    # Ensure RPSessionInterval is set for auto-creation
    try {
        Set-ItemProperty -Path $regPath -Name "RPSessionInterval" -Value 1 -ErrorAction SilentlyContinue
    } catch {}
    Write-Log "Creating a fresh restore point as a health check..."
    try {
        Checkpoint-Computer -Description "PCMasterclass - Verification" -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Log "Restore point created successfully." "INFO"
    } catch {
        # On Windows 11 24H2+ builds, Checkpoint-Computer is broken (srservice removed).
        # System Protection is still working - Windows creates restore points automatically.
        Write-Log "Could not create restore point manually (expected on Win 11 24H2+). System Protection is active." "WARN"
    }
    Write-Log "STATUS: ALREADY_ENABLED" "INFO"
    exit 0
}

# ── Disk Space Assessment ─────────────────────────────────────────────────────
Write-Log "System Protection is currently DISABLED. Assessing disk space..."

$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$totalGB   = [math]::Round($disk.Size / 1GB, 1)
$freeGB    = [math]::Round($disk.FreeSpace / 1GB, 1)
$freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)

Write-Log "Drive C: Total: ${totalGB} GB | Free: ${freeGB} GB (${freePercent}%)"

# Calculate what we'd allocate
$allocPercent = [math]::Min($AllocationPercent, $MaxAllocationPercent)
$allocGB = [math]::Round(($disk.Size * $allocPercent / 100) / 1GB, 1)

Write-Log "Planned allocation for restore points: ${allocGB} GB (${allocPercent}% of drive)"

# ── Space Check ───────────────────────────────────────────────────────────────
$remainingAfterAlloc = $freeGB - $allocGB

if ($freeGB -lt $MinFreeSpaceGB) {
    Write-Log "INSUFFICIENT DISK SPACE to enable restore points." "WARN"
    Write-Log "Free space (${freeGB} GB) is below the ${MinFreeSpaceGB} GB minimum threshold." "WARN"
    Write-Log "This machine may benefit from a storage upgrade or cleanup." "WARN"
    Write-Log "STATUS: LOW_DISK_SPACE | Free: ${freeGB}GB | Total: ${totalGB}GB" "WARN"
    exit 2
}

if ($remainingAfterAlloc -lt ($MinFreeSpaceGB * 0.75)) {
    Write-Log "Enabling restore points would leave only ${remainingAfterAlloc} GB free." "WARN"
    Write-Log "This is tight - machine may benefit from more storage." "WARN"
    Write-Log "Proceeding with a smaller 3% allocation instead." "INFO"
    $allocPercent = 3
    $allocGB = [math]::Round(($disk.Size * $allocPercent / 100) / 1GB, 1)
}

# ── Ensure VSS and related services are ready ────────────────────────────────
Write-Log "Checking VSS and shadow copy services..."
foreach ($svc in @("VSS", "swprv")) {
    try {
        $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($s -and $s.StartType -eq "Disabled") {
            Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue
            Write-Log "Set $svc service to Manual start." "INFO"
        }
    } catch {}
}

# ── Enable System Protection ─────────────────────────────────────────────────
Write-Log "Enabling System Protection on C: drive..."

try {
    Enable-ComputerRestore -Drive $Drive -ErrorAction Stop
    Write-Log "System Protection enabled." "INFO"
} catch {
    Write-Log "Failed to enable System Protection: $_" "ERROR"
    exit 1
}

# Set the allocation size
Write-Log "Setting shadow storage allocation to ${allocPercent}% (${allocGB} GB)..."
try {
    $vssResult = vssadmin resize shadowstorage /for=C: /on=C: /maxsize=${allocPercent}% 2>&1
    Write-Log "Shadow storage configured: $vssResult" "INFO"
} catch {
    Write-Log "Warning - could not set shadow storage size: $_" "WARN"
}

# Ensure the registry has System Restore enabled and auto-create interval set
try {
    Set-ItemProperty -Path $regPath -Name "DisableSR" -Value 0 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regPath -Name "RPSessionInterval" -Value 1 -ErrorAction SilentlyContinue
    Write-Log "Registry: DisableSR=0, RPSessionInterval=1" "INFO"
} catch {}

# ── Create Initial Restore Point ─────────────────────────────────────────────
Write-Log "Creating initial restore point..."
try {
    Checkpoint-Computer -Description "PCMasterclass - Initial Setup" -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
    Write-Log "Initial restore point created successfully." "INFO"
} catch {
    # On Windows 11 24H2+ builds, Checkpoint-Computer is broken (srservice removed).
    # Try WMI fallback.
    Write-Log "Checkpoint-Computer failed, trying WMI fallback..." "WARN"
    try {
        $wmiSR = [wmiclass]"\\.\root\default:SystemRestore"
        $wmiResult = $wmiSR.CreateRestorePoint("PCMasterclass - Initial Setup", 12, 100)
        if ($wmiResult.ReturnValue -eq 0) {
            Write-Log "Initial restore point created via WMI." "INFO"
        } else {
            Write-Log "WMI CreateRestorePoint returned: $($wmiResult.ReturnValue)" "WARN"
        }
    } catch {
        # Both methods failed - likely Windows 11 24H2+ where manual creation is broken.
        # System Protection is still enabled and Windows will create restore points
        # automatically before updates and driver installs.
        Write-Log "Could not create restore point manually. This is expected on Windows 11 24H2+ builds." "WARN"
        Write-Log "System Protection IS enabled - Windows will create restore points automatically before updates." "INFO"
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Log "========================================" "INFO"
Write-Log "System Protection setup complete." "INFO"
Write-Log "  Drive:       C:" "INFO"
Write-Log "  Allocation:  ${allocPercent}% (${allocGB} GB)" "INFO"
Write-Log "  Free after:  ${remainingAfterAlloc} GB" "INFO"
Write-Log "STATUS: ENABLED | Alloc: ${allocGB}GB | Free: ${freeGB}GB | Total: ${totalGB}GB" "INFO"
Write-Log "========================================" "INFO"
exit 0
