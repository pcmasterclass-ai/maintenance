#Requires -RunAsAdministrator
<#
.SYNOPSIS
    PC Masterclass - Automated Maintenance Script
.DESCRIPTION
    Runs a comprehensive maintenance checklist on client Windows PCs and generates
    a branded HTML report with email delivery. Designed to be deployed via TeamViewer
    and run periodically (e.g. 2-3 times per year) or on a scheduled basis.
    Can also be used for ad-hoc health checks.
.NOTES
    Author:  Paul - PC Masterclass
    Version: 2.7.0
    Date:    2026-03-17

    USAGE:
      Run as Administrator (required for SFC, SMART, Windows Update):
        powershell -ExecutionPolicy Bypass -File "PCMasterclass-Maintenance.ps1"

      Optional parameters:
        -SkipSFC         Skip the SFC /SCANNOW check (saves ~15-30 min)
        -SkipUpdates     Skip Windows Update check/install
        -InstallUpdates  Automatically install found updates (default: report only)
        -SkipAdwCleaner  Skip AdwCleaner download and scan
        -SkipUpdate      Skip the automatic update check from GitHub
        -ReportPath      Custom path for the HTML report (default: C:\Teamviewer\Reports)
        -EmailTo         Send the report via email to this address (e.g. paul@pcmasterclass.com.au)
        -SmtpServer      SMTP server for sending email (default: smtp.gmail.com)
        -SmtpPort        SMTP port (default: 587)
        -SmtpUser        SMTP username / email address for authentication
        -SmtpPassword    SMTP password or App Password (Gmail requires App Password)
        -EmailFrom       Sender address (defaults to SmtpUser if not specified)
        -SaveCredential  Save SMTP credentials to an encrypted file for future use
        -CredentialPath  Path to stored encrypted credential (default: C:\Teamviewer\Config\smtp-cred.xml)

    FIRST-TIME SETUP (run once per client machine, as the user account that will run the script):
        powershell -ExecutionPolicy Bypass -File "PCMasterclass-Maintenance.ps1" -SaveCredential `
            -SmtpUser "reports@pcmasterclass.com.au" -SmtpPassword "your-app-password"

    SUBSEQUENT RUNS (credentials loaded automatically from encrypted file):
        powershell -ExecutionPolicy Bypass -File "PCMasterclass-Maintenance.ps1" `
            -EmailTo "paul@pcmasterclass.com.au"

    NOTE ON CREDENTIAL SECURITY:
        Credentials are encrypted using Windows DPAPI, tied to the specific user account
        and machine. The encrypted file cannot be used on another computer or by another
        user. For best security, use a dedicated sending account with a Gmail App Password.
#>

param(
    [switch]$SkipSFC,
    [switch]$SkipUpdates,
    [switch]$InstallUpdates,
    [switch]$SkipAdwCleaner,
    [switch]$SkipUpdate,
    [switch]$Updated,
    [string]$ReportPath = "C:\Teamviewer\Reports",
    [string]$EmailTo = "",
    [string]$SmtpServer = "smtp.gmail.com",
    [int]$SmtpPort = 587,
    [string]$SmtpUser = "",
    [string]$SmtpPassword = "",
    [string]$EmailFrom = "",
    [switch]$SaveCredential,
    [string]$CredentialPath = "C:\Teamviewer\Config\smtp-cred.xml"
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$ScriptVersion = "2.7.0"

# GitHub raw URL for the latest version of this script
# To use: create a private GitHub repo, push the script, and set this URL
# Format: https://raw.githubusercontent.com/OWNER/REPO/main/PCMasterclass-Maintenance.ps1
# For private repos, append ?token=YOUR_PAT or use the $UpdateToken variable below
$UpdateUrl = "https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/PCMasterclass-Maintenance.ps1"
$UpdateToken = ""  # Not needed - repo is public

# ============================================================================
# AUTO-UPDATE FROM GITHUB
# ============================================================================
if (-not $SkipUpdate -and -not $Updated -and -not $SaveCredential) {
    Write-Host "[INFO] Checking for script updates..."
    try {
        $headers = @{}
        if ($UpdateToken) {
            $headers["Authorization"] = "token $UpdateToken"
        }

        # Download the latest version to a temp file
        $tempScript = Join-Path $env:TEMP "PCMasterclass-Maintenance-latest.ps1"
        $webClient = New-Object System.Net.WebClient
        if ($UpdateToken) {
            $webClient.Headers.Add("Authorization", "token $UpdateToken")
        }
        # Force TLS 1.2 for GitHub
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $webClient.DownloadFile($UpdateUrl, $tempScript)

        # Extract version from downloaded script
        $remoteContent = Get-Content $tempScript -Raw
        if ($remoteContent -match '\$ScriptVersion\s*=\s*"([^"]+)"') {
            $remoteVersion = $Matches[1]
        } else {
            $remoteVersion = "0.0.0"
        }

        # Compare versions
        try {
            $localVer = [version]$ScriptVersion
            $remoteVer = [version]$remoteVersion
            $needsUpdate = $remoteVer -gt $localVer
        } catch {
            # Fallback: simple string comparison
            $needsUpdate = $remoteVersion -ne $ScriptVersion
        }

        if ($needsUpdate) {
            Write-Host "[UPDATE] New version available: v$remoteVersion (current: v$ScriptVersion)"
            Write-Host "[UPDATE] Updating script and relaunching..."

            # Overwrite the current script with the new version
            $myPath = $MyInvocation.MyCommand.Path
            Copy-Item -Path $tempScript -Destination $myPath -Force

            # Also update cpu-benchmarks.csv if available on GitHub
            try {
                $csvUrl = $UpdateUrl -replace 'PCMasterclass-Maintenance\.ps1$', 'cpu-benchmarks.csv'
                $csvDest = "C:\Teamviewer\Data\cpu-benchmarks.csv"
                if (-not (Test-Path (Split-Path $csvDest))) { New-Item -Path (Split-Path $csvDest) -ItemType Directory -Force | Out-Null }
                $csvClient = New-Object System.Net.WebClient
                if ($UpdateToken) { $csvClient.Headers.Add("Authorization", "token $UpdateToken") }
                $csvClient.DownloadFile($csvUrl, $csvDest)
                Write-Host "[UPDATE] Updated cpu-benchmarks.csv"
            } catch {
                Write-Host "[WARNING] Could not update cpu-benchmarks.csv: $($_.Exception.Message)" -ForegroundColor Yellow
            }

            # Rebuild the argument list, adding -Updated to prevent loop
            $relaunchArgs = @("-ExecutionPolicy", "Bypass", "-File", "`"$myPath`"", "-Updated")
            if ($SkipSFC) { $relaunchArgs += "-SkipSFC" }
            if ($SkipUpdates) { $relaunchArgs += "-SkipUpdates" }
            if ($InstallUpdates) { $relaunchArgs += "-InstallUpdates" }
            if ($SkipAdwCleaner) { $relaunchArgs += "-SkipAdwCleaner" }
            if ($ReportPath -ne "C:\Teamviewer\Reports") { $relaunchArgs += "-ReportPath"; $relaunchArgs += "`"$ReportPath`"" }
            if ($EmailTo) { $relaunchArgs += "-EmailTo"; $relaunchArgs += "`"$EmailTo`"" }
            if ($SmtpServer -ne "smtp.gmail.com") { $relaunchArgs += "-SmtpServer"; $relaunchArgs += "`"$SmtpServer`"" }
            if ($SmtpPort -ne 587) { $relaunchArgs += "-SmtpPort"; $relaunchArgs += $SmtpPort }
            if ($SmtpUser) { $relaunchArgs += "-SmtpUser"; $relaunchArgs += "`"$SmtpUser`"" }
            if ($EmailFrom) { $relaunchArgs += "-EmailFrom"; $relaunchArgs += "`"$EmailFrom`"" }
            if ($CredentialPath -ne "C:\Teamviewer\Config\smtp-cred.xml") { $relaunchArgs += "-CredentialPath"; $relaunchArgs += "`"$CredentialPath`"" }

            # Clean up temp file and relaunch
            Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
            Start-Process powershell.exe -ArgumentList $relaunchArgs -Wait -NoNewWindow
            exit 0
        } else {
            Write-Host "[INFO] Script is up to date (v$ScriptVersion)"
        }

        # Clean up temp file
        Remove-Item $tempScript -Force -ErrorAction SilentlyContinue

    } catch {
        Write-Host "[WARNING] Auto-update check failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[WARNING] Continuing with current version (v$ScriptVersion)" -ForegroundColor Yellow
    }
}

if ($Updated) {
    Write-Host "[UPDATE] Running updated script v$ScriptVersion"
}
$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ComputerName = $env:COMPUTERNAME
$ReportFile = Join-Path $ReportPath "${ComputerName}_Maintenance_${Timestamp}.html"
$LogFile = Join-Path $ReportPath "${ComputerName}_Maintenance_${Timestamp}.log"

# ============================================================================
# CLIENT NAME LOOKUP (from Rollout Tracker via webhook)
# ============================================================================
$ClientName = ""
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $lookupUrl = "https://script.google.com/macros/s/AKfycbyKkPyodUa3M2Ka9vXzgjMK0hzq6EfA58unifA7Ih6h4OxjLYfXuqea8rrcO2i4yMmF/exec?computerName=$ComputerName&secret=pcm-tracker-2026"
    $lookupResponse = Invoke-RestMethod -Uri $lookupUrl -Method Get -TimeoutSec 15 -ErrorAction Stop
    if ($lookupResponse.status -eq "ok" -and $lookupResponse.clientName) {
        $ClientName = $lookupResponse.clientName
    }
} catch {
    # Non-critical - continue without client name
}

# ============================================================================
# CREDENTIAL MANAGEMENT
# ============================================================================

# Save credentials if requested (one-time setup per machine)
if ($SaveCredential) {
    if (-not $SmtpUser) {
        Write-Host "[ERROR] -SmtpUser is required when saving credentials." -ForegroundColor Red
        exit 1
    }
    if (-not $SmtpPassword) {
        Write-Host "[ERROR] -SmtpPassword is required when saving credentials." -ForegroundColor Red
        exit 1
    }

    # Create config directory with restricted permissions
    $configDir = Split-Path $CredentialPath -Parent
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Restrict folder permissions to current user and SYSTEM only
    try {
        $acl = Get-Acl $configDir
        $acl.SetAccessRuleProtection($true, $false)  # Disable inheritance
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule($currentUser, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.AddAccessRule($userRule)
        $acl.AddAccessRule($systemRule)
        Set-Acl $configDir $acl
    } catch {
        Write-Host "[WARN] Could not restrict folder permissions: $_" -ForegroundColor Yellow
    }

    # Build and save the credential object (encrypted via DPAPI)
    $securePass = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
    $credentialObj = @{
        SmtpUser   = $SmtpUser
        SmtpServer = $SmtpServer
        SmtpPort   = $SmtpPort
        EmailFrom  = if ($EmailFrom) { $EmailFrom } else { $SmtpUser }
        Credential = New-Object System.Management.Automation.PSCredential($SmtpUser, $securePass)
    }
    $credentialObj | Export-Clixml -Path $CredentialPath -Force

    # Also create AES-encrypted fallback credentials (usable by SYSTEM / Scheduled Tasks)
    try {
        $aesKeyPath = Join-Path $configDir "smtp.key"
        $aesConfigPath = Join-Path $configDir "smtp-config.xml"

        # Generate a random 256-bit AES key
        $aesKey = New-Object byte[] 32
        [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($aesKey)
        Set-Content -Path $aesKeyPath -Value $aesKey -Encoding Byte

        # Encrypt the password with AES key (not tied to any user account)
        $securePassAes = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
        $encryptedPass = ConvertFrom-SecureString $securePassAes -Key $aesKey

        $machineConfig = @{
            SmtpUser      = $SmtpUser
            SmtpServer    = $SmtpServer
            SmtpPort      = $SmtpPort
            EmailFrom     = if ($EmailFrom) { $EmailFrom } else { $SmtpUser }
            EncryptedPass = $encryptedPass
        }
        $machineConfig | Export-Clixml -Path $aesConfigPath -Force

        Write-Host ""
        Write-Host "SMTP credentials saved successfully." -ForegroundColor Green
        Write-Host "  DPAPI credential:  $CredentialPath (for interactive use)" -ForegroundColor Gray
        Write-Host "  AES credential:    $aesConfigPath (for SYSTEM / Scheduled Tasks)" -ForegroundColor Gray
        Write-Host "  AES key:           $aesKeyPath" -ForegroundColor Gray
        Write-Host "  User:              $SmtpUser" -ForegroundColor Gray
        Write-Host "  Server:            ${SmtpServer}:${SmtpPort}" -ForegroundColor Gray
        Write-Host ""
    } catch {
        Write-Host ""
        Write-Host "SMTP credentials saved (DPAPI only)." -ForegroundColor Yellow
        Write-Host "  File:      $CredentialPath" -ForegroundColor Gray
        Write-Host "  User:      $SmtpUser" -ForegroundColor Gray
        Write-Host "  Server:    ${SmtpServer}:${SmtpPort}" -ForegroundColor Gray
        Write-Host "  WARNING:   AES fallback creation failed: $_" -ForegroundColor Yellow
        Write-Host "  Note:      Scheduled tasks running as SYSTEM may not be able to send email." -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "You can now run the script without -SmtpUser/-SmtpPassword:" -ForegroundColor Cyan
    Write-Host "  .\PCMasterclass-Maintenance.ps1 -EmailTo paul@pcmasterclass.com.au" -ForegroundColor Cyan
    Write-Host ""
    exit 0
}

# Load saved credentials if no SMTP credentials provided on command line
$StoredCredential = $null
if ($EmailTo -and (-not $SmtpUser -or -not $SmtpPassword)) {

    $credLoaded = $false

    # Attempt 1: DPAPI-encrypted credential (works when running as the user who saved it)
    if (Test-Path $CredentialPath) {
        try {
            $StoredCredential = Import-Clixml -Path $CredentialPath
            # Test that we can actually read the password (will fail under different user/SYSTEM)
            $testPass = $StoredCredential.Credential.GetNetworkCredential().Password
            if ($testPass) {
                $SmtpUser   = $StoredCredential.SmtpUser
                $SmtpServer = if ($StoredCredential.SmtpServer) { $StoredCredential.SmtpServer } else { $SmtpServer }
                $SmtpPort   = if ($StoredCredential.SmtpPort)   { $StoredCredential.SmtpPort }   else { $SmtpPort }
                if (-not $EmailFrom) { $EmailFrom = $StoredCredential.EmailFrom }
                $credLoaded = $true
                $testPass = $null
            }
        } catch {
            Write-Host "[INFO] DPAPI credential not usable in this context (likely running as SYSTEM). Trying fallback..." -ForegroundColor Gray
            $StoredCredential = $null
        }
    }

    # Attempt 2: AES-encrypted fallback (works when running as SYSTEM via Scheduled Task)
    if (-not $credLoaded) {
        $configDir = Split-Path $CredentialPath -Parent
        $aesKeyPath = Join-Path $configDir "smtp.key"
        $aesConfigPath = Join-Path $configDir "smtp-config.xml"

        if ((Test-Path $aesKeyPath) -and (Test-Path $aesConfigPath)) {
            try {
                $aesKey = Get-Content $aesKeyPath -Encoding Byte
                $machineConfig = Import-Clixml -Path $aesConfigPath
                $secPass = ConvertTo-SecureString $machineConfig.EncryptedPass -Key $aesKey

                $StoredCredential = @{
                    SmtpUser   = $machineConfig.SmtpUser
                    SmtpServer = $machineConfig.SmtpServer
                    SmtpPort   = $machineConfig.SmtpPort
                    EmailFrom  = $machineConfig.EmailFrom
                    Credential = New-Object System.Management.Automation.PSCredential($machineConfig.SmtpUser, $secPass)
                }

                $SmtpUser   = $StoredCredential.SmtpUser
                $SmtpServer = if ($StoredCredential.SmtpServer) { $StoredCredential.SmtpServer } else { $SmtpServer }
                $SmtpPort   = if ($StoredCredential.SmtpPort)   { $StoredCredential.SmtpPort }   else { $SmtpPort }
                if (-not $EmailFrom) { $EmailFrom = $StoredCredential.EmailFrom }
                $credLoaded = $true

                Write-Host "[INFO] Loaded AES-encrypted fallback credentials for $SmtpUser" -ForegroundColor Gray
            } catch {
                Write-Host "[ERROR] Failed to load AES fallback credentials: $_" -ForegroundColor Red
            }
        }
    }

    if (-not $credLoaded) {
        Write-Host "[WARN] Email requested but no usable credentials found." -ForegroundColor Yellow
        Write-Host "Run with -SaveCredential first, or provide -SmtpUser and -SmtpPassword." -ForegroundColor Yellow
    }
}

# Suspicious startup program keywords  - add to this list as needed
$SuspiciousKeywords = @(
    "anydesk", "rustdesk", "ammyy", "supremo", "ultraviewer", "remotepc",
    "logmein", "gotomypc", "splashtop", "connectwise", "screenconnect",
    "vnc", "tightvnc", "realvnc", "radmin", "dameware",
    "update", "helper", "service", "loader", "runtime", "host",
    "temp", "tmp", "appdata"
)

# Known-safe startup programs (won't be flagged even if they match generic names)
$KnownSafePrograms = @(
    # Windows and Microsoft
    "SecurityHealth", "Windows Security", "OneDrive", "Microsoft Teams",
    "MicrosoftEdgeUpdate", "Office Automatic Updates", "Office Feature Updates",
    "DirectXDatabaseUpdater", "SystemSoundsService", "MsCtfMonitor",
    "KeyPreGenTask", "UserTask", "Calibration Loader",
    "AD RMS Rights Policy",
    # Common legitimate software updaters
    "GoogleUpdater", "DropboxUpdater", "Adobe", "Zoom",
    # Hardware drivers
    "RealTek", "NVIDIA", "Intel", "Synaptics", "AMD",
    # Common user apps
    "iTunes Helper", "Spotify", "Steam", "Discord",
    # Our tools
    "TeamViewer", "PCMasterclass"
)

# ============================================================================
# INITIALISATION
# ============================================================================

# Ensure report directory exists
if (-not (Test-Path $ReportPath)) {
    New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
}

# Results collection  - each module adds its results here
$Results = [ordered]@{
    ScriptVersion    = $ScriptVersion
    ComputerName     = $ComputerName
    RunTimestamp     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    OSVersion        = ""
    SystemInfo       = @{}
    SFC              = @{}
    DiskHealth       = @{}
    WindowsUpdates   = @{}
    iDriveBackup     = @{}
    Malwarebytes     = @{}
    Defender         = @{}
    EventLogErrors   = @{}
    PendingReboot    = @{}
    Firewall         = @{}
    UserAccounts     = @{}
    DISM             = @{}
    TempFiles        = @{}
    StartupPrograms  = @{}
    ScheduledTasks   = @{}
    BrowserExtensions = @{}
    ServiceStatus    = @{}
    NetworkConfig    = @{}
    AdwCleaner       = @{}
    RestorePoints    = @{}
    TelemetryServices = @{}
    Errors           = @()
    EmailResult      = @{}
    WebhookResult    = @{}
}

# Script timer
$scriptStartTime = Get-Date

# Logging helper
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "[$Level] $(Get-Date -Format 'HH:mm:ss') - $Message"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry
}

Write-Log "============================================"
Write-Log "PC Masterclass Maintenance Script v$ScriptVersion"
Write-Log "Computer: $ComputerName"
Write-Log "============================================"

# ============================================================================
# MODULE 0: SYSTEM INFO
# ============================================================================
Write-Log "Gathering system information..."

try {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $Results.OSVersion = "$($os.Caption) - Build $($os.BuildNumber)"

    # CPU info
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $cpuName = ($cpu.Name -replace '\s+', ' ').Trim()
    $cpuCores = $cpu.NumberOfCores
    $cpuThreads = $cpu.NumberOfLogicalProcessors
    $cpuMaxSpeed = [math]::Round($cpu.MaxClockSpeed / 1000, 2)

    # Serial number / service tag
    $bios = Get-CimInstance Win32_BIOS
    $serialNumber = $bios.SerialNumber
    if (-not $serialNumber -or $serialNumber -eq "To Be Filled By O.E.M." -or $serialNumber -eq "Default string") {
        $serialNumber = "Not available"
    }

    # BitLocker status (may require admin / may not be available on Home editions)
    $bitlockerStatus = @()
    try {
        $blVolumes = Get-BitLockerVolume -ErrorAction Stop
        foreach ($blVol in $blVolumes) {
            $bitlockerStatus += @{
                Drive           = $blVol.MountPoint
                ProtectionStatus = $blVol.ProtectionStatus.ToString()
                EncryptionMethod = $blVol.EncryptionMethod.ToString()
                VolumeStatus    = $blVol.VolumeStatus.ToString()
            }
        }
    } catch {
        # BitLocker cmdlet not available (e.g. Home edition) or access denied
        $bitlockerStatus += @{
            Drive           = "N/A"
            ProtectionStatus = "Not available (Home edition or access denied)"
            EncryptionMethod = "N/A"
            VolumeStatus    = "N/A"
        }
    }

    # Network adapters (active only - with IP addresses)
    $networkAdapters = @()
    try {
        $adapters = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
        foreach ($adapter in $adapters) {
            $ipv4 = ($adapter.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ", "
            $gateway = if ($adapter.DefaultIPGateway) { $adapter.DefaultIPGateway -join ", " } else { "None" }
            $dns = if ($adapter.DNSServerSearchOrder) { $adapter.DNSServerSearchOrder -join ", " } else { "None" }
            $mac = $adapter.MACAddress
            $networkAdapters += @{
                Name    = $adapter.Description
                IPv4    = $ipv4
                Gateway = $gateway
                DNS     = $dns
                MAC     = $mac
                DHCP    = $adapter.DHCPEnabled
            }
        }
    } catch {
        $networkAdapters += @{ Name = "Error retrieving network info"; IPv4 = "N/A"; Gateway = "N/A"; DNS = "N/A"; MAC = "N/A"; DHCP = $false }
    }

    $Results.SystemInfo = @{
        OS              = $os.Caption
        Build           = $os.BuildNumber
        LastBoot        = $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
        UptimeDays      = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
        TotalRAM_GB     = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        Manufacturer    = $cs.Manufacturer
        Model           = $cs.Model
        Domain          = $cs.Domain
        CPU             = $cpuName
        CPUCores        = $cpuCores
        CPUThreads      = $cpuThreads
        CPUMaxSpeedGHz  = $cpuMaxSpeed
        SerialNumber    = $serialNumber
        BIOSVersion     = $bios.SMBIOSBIOSVersion
        BitLocker       = $bitlockerStatus
        NetworkAdapters = $networkAdapters
        Temperatures    = @()
    }

    # Thermal zone temperatures
    # Method 1: MSAcpi_ThermalZoneTemperature (most common on desktops/laptops)
    $thermalData = @()
    try {
        $thermalZones = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop
        foreach ($tz in $thermalZones) {
            # WMI returns temperature in tenths of Kelvin - convert to Celsius
            $tempC = [math]::Round(($tz.CurrentTemperature / 10) - 273.15, 1)
            # Discard obviously invalid readings (sensor errors return near-zero Kelvin or extreme values)
            if ($tempC -lt -40 -or $tempC -gt 150) { continue }
            $zoneName = if ($tz.InstanceName) { ($tz.InstanceName -split '\\')[-1] -replace '_', ' ' } else { "Thermal Zone" }
            $thermalData += @{ Zone = $zoneName; TempC = $tempC; TempF = [math]::Round(($tempC * 9/5) + 32, 1) }
        }
    } catch {
        # MSAcpi not available - try alternative
    }

    # Method 2: Win32_TemperatureProbe (less common but worth trying)
    if ($thermalData.Count -eq 0) {
        try {
            $tempProbes = Get-CimInstance Win32_TemperatureProbe -ErrorAction Stop | Where-Object { $_.CurrentReading }
            foreach ($probe in $tempProbes) {
                $tempC = [math]::Round($probe.CurrentReading / 10, 1)
                if ($tempC -lt -40 -or $tempC -gt 150) { continue }
                $probeName = if ($probe.Description) { $probe.Description } else { "Temperature Probe" }
                $thermalData += @{ Zone = $probeName; TempC = $tempC; TempF = [math]::Round(($tempC * 9/5) + 32, 1) }
            }
        } catch {
            # Not available
        }
    }

    # Method 3: Performance counters for thermal zone info
    if ($thermalData.Count -eq 0) {
        try {
            $perfThermal = Get-CimInstance Win32_PerfFormattedData_Counters_ThermalZoneInformation -ErrorAction Stop
            foreach ($pt in $perfThermal) {
                $tempC = [math]::Round($pt.Temperature - 273.15, 1)
                if ($tempC -lt -40 -or $tempC -gt 150) { continue }
                $zoneName = if ($pt.Name) { $pt.Name } else { "Thermal Zone" }
                $thermalData += @{ Zone = $zoneName; TempC = $tempC; TempF = [math]::Round(($tempC * 9/5) + 32, 1) }
            }
        } catch {
            # Not available
        }
    }

    if ($thermalData.Count -eq 0) {
        $thermalData += @{ Zone = "N/A"; TempC = 0; TempF = 0; Note = "Temperature sensors not accessible via WMI" }
    }

    $Results.SystemInfo.Temperatures = $thermalData

    Write-Log "OS: $($Results.OSVersion) | Uptime: $($Results.SystemInfo.UptimeDays) days"
    Write-Log "CPU: $cpuName ($cpuCores cores / $cpuThreads threads @ ${cpuMaxSpeed}GHz)"
    Write-Log "RAM: $($Results.SystemInfo.TotalRAM_GB)GB | $($cs.Manufacturer) $($cs.Model) | S/N: $serialNumber"
    if ($thermalData[0].Zone -ne "N/A") {
        $tempSummary = ($thermalData | ForEach-Object { "$($_.Zone): $($_.TempC)C" }) -join " | "
        Write-Log "Temps: $tempSummary"
    } else {
        Write-Log "Temps: Not available via WMI"
    }
} catch {
    Write-Log "Failed to gather system info: $_" "ERROR"
    $Results.Errors += "System Info: $_"
}

# ============================================================================
# CPU BENCHMARK LOOKUP (PassMark)
# Looks up the detected CPU in the bundled benchmark CSV to provide a
# performance score and colour-coded tier rating in the report.
# ============================================================================
Write-Log "Looking up CPU benchmark..."

try {
    $benchmarkCsvPath = "C:\Teamviewer\Data\cpu-benchmarks.csv"
    $Results.SystemInfo.CPUBenchmark = @{
        Score          = 0
        Tier           = "Unknown"
        TierColor      = "#6c757d"
        Matched        = ""
        DefinitionsDate = ""
        DefinitionsAge  = ""
    }

    if (Test-Path $benchmarkCsvPath) {
        $benchmarkData = Import-Csv $benchmarkCsvPath

        # Track definitions age from file last-write time
        $csvDate = (Get-Item $benchmarkCsvPath).LastWriteTime
        $Results.SystemInfo.CPUBenchmark.DefinitionsDate = $csvDate.ToString("d MMM yyyy")
        $ageDays = [math]::Floor(((Get-Date) - $csvDate).TotalDays)
        if ($ageDays -eq 0) {
            $Results.SystemInfo.CPUBenchmark.DefinitionsAge = "today"
        } elseif ($ageDays -eq 1) {
            $Results.SystemInfo.CPUBenchmark.DefinitionsAge = "1 day ago"
        } elseif ($ageDays -lt 30) {
            $Results.SystemInfo.CPUBenchmark.DefinitionsAge = "$ageDays days ago"
        } elseif ($ageDays -lt 365) {
            $months = [math]::Floor($ageDays / 30)
            $Results.SystemInfo.CPUBenchmark.DefinitionsAge = "$months month$(if ($months -gt 1) {'s'}) ago"
        } else {
            $years = [math]::Floor($ageDays / 365)
            $Results.SystemInfo.CPUBenchmark.DefinitionsAge = "$years year$(if ($years -gt 1) {'s'}) ago"
        }
        $detectedCpu = $Results.SystemInfo.CPU

        # Try exact match first (after normalising whitespace)
        $normalised = ($detectedCpu -replace '\s+', ' ').Trim()
        $match = $benchmarkData | Where-Object { $_.CpuName -eq $normalised } | Select-Object -First 1

        # Try match without clock speed suffix (e.g. "@ 3.60GHz")
        if (-not $match) {
            $cpuNoSpeed = ($normalised -replace '\s*@\s*[\d\.]+\s*GHz', '').Trim()
            $match = $benchmarkData | Where-Object { $_.CpuName -eq $cpuNoSpeed } | Select-Object -First 1
        }

        # Try match adding clock speed from WMI
        if (-not $match -and $cpuMaxSpeed) {
            $cpuWithSpeed = "$cpuNoSpeed @ $($cpuMaxSpeed.ToString('0.00'))GHz"
            $match = $benchmarkData | Where-Object { $_.CpuName -eq $cpuWithSpeed } | Select-Object -First 1
        }

        # Try fuzzy match - find closest match by checking if CSV name is contained in detected name or vice versa
        if (-not $match) {
            $match = $benchmarkData | Where-Object {
                $csvName = ($_.CpuName -replace '\s*@\s*[\d\.]+\s*GHz', '').Trim()
                $cpuNoSpeed -like "*$csvName*" -or $csvName -like "*$cpuNoSpeed*"
            } | Select-Object -First 1
        }

        # Try partial model match (e.g. "i7-14700" from "Intel Core i7-14700")
        if (-not $match) {
            # Extract the model number pattern like "i7-14700" or "Ryzen 7 5800X"
            if ($cpuNoSpeed -match '(i[3579]-\d{4,5}\w*|Ryzen\s+\d\s+\d{4}\w*|N\d{2,3}|Celeron\s+\w+\d+|Pentium\s+\w+\s+\w+\d+)') {
                $modelPattern = $Matches[1]
                $match = $benchmarkData | Where-Object {
                    $_.CpuName -match [regex]::Escape($modelPattern)
                } | Select-Object -First 1
            }
        }

        if ($match) {
            $score = [int]$match.CpuMark
            $Results.SystemInfo.CPUBenchmark.Score = $score
            $Results.SystemInfo.CPUBenchmark.Matched = $match.CpuName

            # Performance tiers based on average consumer desktop CPU (~15,000 in 2025/2026)
            # Very Slow: < 5,000     (old dual-core, Celerons, 4th gen and below)
            # Slow:      5,000-9,999 (budget / older i3/i5, entry Ryzen)
            # Average:   10,000-19,999 (mainstream i5 10th-11th gen, Ryzen 5 3xxx)
            # Fast:      20,000-34,999 (current-gen i5/i7, Ryzen 5/7 5xxx-7xxx)
            # Very Fast: 35,000+      (high-end i7/i9, Ryzen 7/9 latest gen)
            if ($score -ge 35000) {
                $Results.SystemInfo.CPUBenchmark.Tier = "Very Fast"
                $Results.SystemInfo.CPUBenchmark.TierColor = "#28a745"  # Green
            } elseif ($score -ge 20000) {
                $Results.SystemInfo.CPUBenchmark.Tier = "Fast"
                $Results.SystemInfo.CPUBenchmark.TierColor = "#17a2b8"  # Blue
            } elseif ($score -ge 10000) {
                $Results.SystemInfo.CPUBenchmark.Tier = "Average"
                $Results.SystemInfo.CPUBenchmark.TierColor = "#ffc107"  # Yellow/Amber
            } elseif ($score -ge 5000) {
                $Results.SystemInfo.CPUBenchmark.Tier = "Slow"
                $Results.SystemInfo.CPUBenchmark.TierColor = "#fd7e14"  # Orange
            } else {
                $Results.SystemInfo.CPUBenchmark.Tier = "Very Slow"
                $Results.SystemInfo.CPUBenchmark.TierColor = "#dc3545"  # Red
            }

            Write-Log "CPU Benchmark: $score (PassMark) - $($Results.SystemInfo.CPUBenchmark.Tier) [matched: $($match.CpuName)]"
        } else {
            Write-Log "CPU Benchmark: No match found for '$detectedCpu'" "WARN"
        }
    } else {
        Write-Log "CPU Benchmark: csv not found at $benchmarkCsvPath" "WARN"
    }
} catch {
    Write-Log "CPU Benchmark lookup failed: $_" "WARN"
}


# ============================================================================
# MODULE 1: DISM COMPONENT STORE HEALTH (with auto-repair)
# Runs BEFORE SFC so that if the component store is repaired,
# SFC can verify system files against the healthy store.
# ============================================================================
Write-Log "Checking DISM component store health..."

try {
    $dismStatus = "PASS"
    $dismResult = "Healthy"
    $dismRepaired = $false
    $dismRepairOutput = ""

    # Step 1: Check health
    Write-Log "Running DISM /CheckHealth..."
    $dismOutput = DISM /Online /Cleanup-Image /CheckHealth 2>&1 | Out-String

    if ($dismOutput -match "The component store is repairable") {
        # Step 2: Auto-repair if repairable (no restart required)
        Write-Log "Component store is repairable - running RestoreHealth (this may take 15-30 minutes)..."
        $dismRepairOutput = DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String

        if ($dismRepairOutput -match "The restore operation completed successfully") {
            $dismStatus = "PASS - Repaired"
            $dismResult = "Component store was repairable and has been repaired successfully"
            $dismRepaired = $true
            Write-Log "DISM repair completed successfully"
        } elseif ($dismRepairOutput -match "Error") {
            $dismStatus = "WARNING - Repair attempted but may not have completed"
            $dismResult = "Repair was attempted but encountered errors - manual review recommended"
            Write-Log "DISM repair encountered errors" "WARN"
        } else {
            $dismStatus = "PASS - Repaired"
            $dismResult = "Component store repair completed"
            $dismRepaired = $true
            Write-Log "DISM repair completed"
        }
    } elseif ($dismOutput -match "The component store is not repairable") {
        $dismStatus = "FAIL - Component store corruption detected"
        $dismResult = "Not repairable by DISM - may need Windows repair or reinstall"
    } elseif ($dismOutput -match "No component store corruption detected") {
        $dismStatus = "PASS"
        $dismResult = "No corruption detected"
    } elseif ($dismOutput -match "healthy") {
        $dismStatus = "PASS"
        $dismResult = "Healthy"
    }

    $Results.DISM = @{
        Status   = $dismStatus
        Result   = $dismResult
        Repaired = $dismRepaired
        Output   = if ($dismOutput.Length -gt 500) { $dismOutput.Substring(0, 500) + "..." } else { $dismOutput }
    }

    Write-Log "DISM: $dismResult"

} catch {
    Write-Log "DISM check failed: $_" "ERROR"
    $Results.DISM = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "DISM: $_"
}


# ============================================================================
# MODULE 2: SFC /SCANNOW
# ============================================================================
if (-not $SkipSFC) {
    Write-Log "Starting SFC /SCANNOW (this may take 15-30 minutes)..."
    $sfcStart = Get-Date

    try {
        $sfcOutput = & sfc /scannow 2>&1 | Out-String
        $sfcDuration = [math]::Round(((Get-Date) - $sfcStart).TotalMinutes, 1)

        # Strip null bytes (SFC sometimes outputs UTF-16 with embedded nulls)
        $sfcClean = $sfcOutput -replace '\x00', ''

        # Parse the SFC result
        $sfcStatus = "Unknown"
        if ($sfcClean -match "did not find any integrity violations") {
            $sfcStatus = "PASS"
        } elseif ($sfcClean -match "found corrupt files and successfully repaired") {
            $sfcStatus = "REPAIRED"
        } elseif ($sfcClean -match "found corrupt files but was unable to fix") {
            $sfcStatus = "FAIL"
        } elseif ($sfcClean -match "could not perform the requested operation") {
            $sfcStatus = "ERROR"
        }

        # Also check the CBS log for details
        $cbsLogPath = "$env:SystemRoot\Logs\CBS\CBS.log"
        $cbsCorruptCount = 0
        if (Test-Path $cbsLogPath) {
            $cbsRecent = Get-Content $cbsLogPath -Tail 200 | Where-Object { $_ -match "corrupt|Cannot repair" }
            $cbsCorruptCount = ($cbsRecent | Measure-Object).Count
        }

        $Results.SFC = @{
            Status          = $sfcStatus
            DurationMinutes = $sfcDuration
            CorruptEntries  = $cbsCorruptCount
            RawOutput       = $sfcOutput.Trim()
        }
        Write-Log "SFC completed in ${sfcDuration}m  - Status: $sfcStatus"

    } catch {
        Write-Log "SFC failed: $_" "ERROR"
        $Results.SFC = @{ Status = "ERROR"; Error = $_.ToString() }
        $Results.Errors += "SFC: $_"
    }
} else {
    Write-Log "SFC /SCANNOW skipped (SkipSFC flag set)"
    $Results.SFC = @{ Status = "SKIPPED" }
}


# ============================================================================
# MODULE 3: DISK HEALTH (SMART)
# ============================================================================
Write-Log "Checking disk health..."

try {
    $disks = Get-PhysicalDisk | Select-Object DeviceId, FriendlyName, MediaType, Size, HealthStatus, OperationalStatus

    $diskResults = @()
    foreach ($disk in $disks) {
        $sizeGB = [math]::Round($disk.Size / 1GB, 1)

        # Get reliability counters if available
        $reliability = $null
        try {
            $reliability = Get-PhysicalDisk -UniqueId $disk.UniqueId | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
        } catch { }

        $diskInfo = @{
            Name              = $disk.FriendlyName
            Type              = $disk.MediaType
            SizeGB            = $sizeGB
            HealthStatus      = $disk.HealthStatus.ToString()
            OperationalStatus = $disk.OperationalStatus.ToString()
            Temperature       = if ($reliability) { $reliability.Temperature } else { "N/A" }
            ReadErrors        = if ($reliability) { $reliability.ReadErrorsTotal } else { "N/A" }
            WearLevel         = if ($reliability) { $reliability.Wear } else { "N/A" }
        }
        $diskResults += $diskInfo

        Write-Log "Disk: $($disk.FriendlyName) ($($disk.MediaType))  - $sizeGB GB  - Health: $($disk.HealthStatus)"
    }

    # Also check disk space on all volumes
    $volumes = Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq 'Fixed' } |
        Select-Object DriveLetter, FileSystemLabel, @{N='SizeGB';E={[math]::Round($_.Size/1GB,1)}},
            @{N='FreeGB';E={[math]::Round($_.SizeRemaining/1GB,1)}},
            @{N='FreePercent';E={[math]::Round(($_.SizeRemaining/$_.Size)*100,1)}}

    $volumeResults = @()
    foreach ($vol in $volumes) {
        $volumeResults += @{
            Drive       = "$($vol.DriveLetter):"
            Label       = $vol.FileSystemLabel
            SizeGB      = $vol.SizeGB
            FreeGB      = $vol.FreeGB
            FreePercent = $vol.FreePercent
            Warning     = ($vol.FreePercent -lt 10)
        }
        $freeWarning = if ($vol.FreePercent -lt 10) { " *** LOW SPACE WARNING ***" } else { "" }
        Write-Log "  Volume $($vol.DriveLetter): $($vol.FreeGB)GB free of $($vol.SizeGB)GB ($($vol.FreePercent)%)$freeWarning"
    }

    $Results.DiskHealth = @{
        Status   = if ($disks | Where-Object { $_.HealthStatus -ne "Healthy" }) { "WARNING" } else { "PASS" }
        Disks    = $diskResults
        Volumes  = $volumeResults
    }

} catch {
    Write-Log "Disk health check failed: $_" "ERROR"
    $Results.DiskHealth = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Disk Health: $_"
}


# ============================================================================
# MODULE 4: WINDOWS UPDATES
# ============================================================================
if (-not $SkipUpdates) {
    Write-Log "Checking for Windows Updates..."

    try {
        # Install PSWindowsUpdate module if not present
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log "Installing PSWindowsUpdate module..."
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
            Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop

        # Check for available updates
        Write-Log "Scanning for available updates (this may take a few minutes)..."
        $updates = Get-WindowsUpdate -ErrorAction Stop

        $updateList = @()
        foreach ($update in $updates) {
            $updateList += @{
                Title    = $update.Title
                KB       = $update.KBArticleIDs -join ", "
                SizeMB   = [math]::Round($update.MaxDownloadSize / 1MB, 1)
                IsImportant = $update.MsrcSeverity -in @("Critical", "Important")
            }
        }

        $pendingCount = ($updates | Measure-Object).Count
        Write-Log "Found $pendingCount pending update(s)"

        # Install updates if flag is set
        $installResult = "Not attempted"
        if ($InstallUpdates -and $pendingCount -gt 0) {
            Write-Log "Installing updates (InstallUpdates flag set)..."
            $installed = Get-WindowsUpdate -Install -AcceptAll -IgnoreReboot -ErrorAction Stop
            $installResult = "Installed $($installed.Count) update(s)"
            Write-Log $installResult

            # Check if reboot is needed
            $rebootPending = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) -ne $null
            if ($rebootPending) {
                $installResult += "  - REBOOT REQUIRED"
                Write-Log "Reboot is required to complete updates" "WARN"
            }
        }

        $Results.WindowsUpdates = @{
            Status         = if ($pendingCount -eq 0) { "UP TO DATE" } elseif ($pendingCount -le 3) { "UPDATES AVAILABLE" } else { "WARNING" }
            PendingCount   = $pendingCount
            Updates        = $updateList
            InstallResult  = $installResult
        }

    } catch {
        Write-Log "Windows Update check failed: $_" "ERROR"
        $Results.WindowsUpdates = @{ Status = "ERROR"; Error = $_.ToString() }
        $Results.Errors += "Windows Updates: $_"
    }
} else {
    Write-Log "Windows Update check skipped (SkipUpdates flag set)"
    $Results.WindowsUpdates = @{ Status = "SKIPPED" }
}


# ============================================================================
# MODULE 5: iDRIVE BACKUP VERIFICATION
# ============================================================================
Write-Log "Checking iDrive backup status..."

try {
    $idriveStatus = "NOT INSTALLED"
    $lastBackupInfo = "Could not determine"
    $lastBackupDate = "Unknown"
    $backupFilesCount = "N/A"
    $filesConsidered = "N/A"
    $backupDuration = "N/A"
    $backupComputerName = "N/A"
    $backupAccount = "N/A"
    $filesInSync = "N/A"

    # Check if iDrive service is running
    $idriveService = Get-Service -Name "IDrive*" -ErrorAction SilentlyContinue
    $idriveServiceRunning = ($idriveService | Where-Object { $_.Status -eq "Running" }).Count -gt 0

    # Check for iDrive install
    $idriveInstalled = $false
    $installPaths = @(
        "$env:ProgramFiles\IDriveWindows",
        "${env:ProgramFiles(x86)}\IDriveWindows",
        "C:\IDriveWindows"
    )
    foreach ($iPath in $installPaths) {
        if (Test-Path $iPath) { $idriveInstalled = $true; break }
    }
    if ($idriveService) { $idriveInstalled = $true }

    if ($idriveInstalled) {
        $idriveStatus = "INSTALLED"

        # Primary method: Parse iDrive Session LOGXML files
        # These are daily XML logs at C:\ProgramData\IDrive\IBCOMMON\<profile>\Session\LOGXML\
        $logXmlFound = $false
        $idriveDataRoot = "$env:ProgramData\IDrive"

        if (Test-Path $idriveDataRoot) {
            # Find LOGXML directories (search under IBCOMMON\*\Session\LOGXML)
            $logXmlDirs = Get-ChildItem -LiteralPath "$idriveDataRoot\IBCOMMON" -Directory -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $sessionLogXml = Join-Path $_.FullName "Session\LOGXML"
                    if (Test-Path $sessionLogXml) { $sessionLogXml }
                }

            foreach ($logXmlDir in $logXmlDirs) {
                # Get the most recent XML log file
                $recentXml = Get-ChildItem -LiteralPath $logXmlDir -Filter "*.xml" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1

                if ($recentXml) {
                    $logXmlFound = $true
                    $xmlContent = Get-Content -LiteralPath $recentXml.FullName -ErrorAction SilentlyContinue | Out-String

                    if ($xmlContent) {
                        # Parse the XML
                        try {
                            [xml]$logXml = $xmlContent
                            $record = $logXml.records.record

                            if ($record) {
                                # Handle multiple records - get the last one
                                if ($record -is [System.Array]) {
                                    $record = $record[-1]
                                }

                                $backupStatus = $record.status
                                $backupDateTime = $record.DateTime
                                $backupFilesCount = $record.bkpfiles
                                $filesConsidered = $record.files
                                $filesInSync = $record.filesinsync
                                $backupDuration = $record.duration
                                $backupComputerName = $record.mpc
                                $backupAccount = $record.uname

                                # Parse the backup date
                                if ($backupDateTime) {
                                    $lastBackupDate = $backupDateTime
                                    # Try to calculate days since backup
                                    $parsedDate = $null
                                    $dateFormats = @("MM-dd-yyyy HH:mm:ss", "dd-MM-yyyy HH:mm:ss", "yyyy-MM-dd HH:mm:ss")
                                    foreach ($fmt in $dateFormats) {
                                        $parsedDate = [DateTime]::ParseExact($backupDateTime, $fmt, $null) 2>$null
                                        if ($parsedDate) { break }
                                    }
                                    if (-not $parsedDate) {
                                        try { $parsedDate = [DateTime]::Parse($backupDateTime) } catch {}
                                    }
                                }

                                # Calculate file counts for detail reporting
                                $filesTotal = 0
                                $filesBacked = 0
                                $filesFailed = 0
                                try {
                                    if ($filesConsidered) { $filesTotal = [int]$filesConsidered }
                                    if ($backupFilesCount) { $filesBacked = [int]$backupFilesCount }
                                    $filesFailed = $filesTotal - $filesBacked
                                    if ($filesFailed -lt 0) { $filesFailed = 0 }
                                } catch { $filesFailed = 0 }

                                # Determine status
                                # "Success*" means completed with some files skipped (e.g. Outlook PST locked)
                                # This is normal and should not be treated as a hard failure
                                if ($backupStatus -match "^Success") {
                                    if ($filesFailed -gt 0) {
                                        $lastBackupInfo = "Completed - $filesFailed of $filesTotal item(s) could not be backed up (e.g. files in use)"
                                    } else {
                                        $lastBackupInfo = "Success - all $filesTotal files backed up"
                                    }
                                    if ($parsedDate) {
                                        $daysSinceBackup = [math]::Round(((Get-Date) - $parsedDate).TotalDays, 1)
                                        if ($daysSinceBackup -le 2) {
                                            $idriveStatus = "PASS"
                                        } else {
                                            $idriveStatus = "WARNING - Last backup $daysSinceBackup days ago"
                                        }
                                    } else {
                                        $idriveStatus = "PASS"
                                    }
                                } elseif ($backupStatus -match "Failure|Error|Failed") {
                                    if ($filesFailed -gt 0 -and $filesTotal -gt 0) {
                                        $lastBackupInfo = "Failed - $filesFailed of $filesTotal item(s) could not be backed up"
                                    } else {
                                        $lastBackupInfo = "Failed - check iDrive logs for details"
                                    }
                                    # If most files succeeded, treat as warning not hard fail
                                    if ($filesTotal -gt 0 -and $filesBacked -gt 0 -and ($filesBacked / $filesTotal) -gt 0.95) {
                                        $idriveStatus = "WARNING - Partial backup failure"
                                    } else {
                                        $idriveStatus = "FAIL"
                                    }
                                } elseif ($backupStatus -match "Missed") {
                                    $lastBackupInfo = "Missed - scheduled backup did not run"
                                    $idriveStatus = "WARNING - Backup missed"
                                } else {
                                    $lastBackupInfo = "Status: $backupStatus"
                                    $idriveStatus = "CHECK REQUIRED"
                                }
                            }
                        } catch {
                            Write-Log "Could not parse iDrive XML log: $_" "WARN"
                        }
                    }
                    # Only need one LOGXML directory
                    break
                }
            }
        }

        # Fallback: Check the daily logs directory if LOGXML not found
        if (-not $logXmlFound -and (Test-Path $idriveDataRoot)) {
            $dailyLogDirs = Get-ChildItem -LiteralPath "$idriveDataRoot\IBCOMMON\logs" -Directory -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.GetFiles("*.xml", [System.IO.SearchOption]::TopDirectoryOnly).Count -gt 0 }

            foreach ($dld in $dailyLogDirs) {
                $recentDailyXml = Get-ChildItem -LiteralPath $dld.FullName -Filter "*.xml" -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "^\d{2}-\d{2}-\d{4}\.xml$" } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1

                if ($recentDailyXml) {
                    $lastBackupDate = $recentDailyXml.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $daysSinceBackup = [math]::Round(((Get-Date) - $recentDailyXml.LastWriteTime).TotalDays, 1)
                    $lastBackupInfo = "Log found (daily log)"
                    if ($daysSinceBackup -le 2) {
                        $idriveStatus = "PASS"
                    } else {
                        $idriveStatus = "WARNING - Last log $daysSinceBackup days ago"
                    }
                    break
                }
            }
        }

        $Results.iDriveBackup = @{
            Status           = $idriveStatus
            Installed        = $true
            ServiceRunning   = $idriveServiceRunning
            LastBackupDate   = $lastBackupDate
            LastBackupResult = $lastBackupInfo
            FilesBackedUp    = $backupFilesCount
            FilesTotal       = if ($filesConsidered) { $filesConsidered } else { "N/A" }
            FilesInSync      = $filesInSync
            Duration         = $backupDuration
            ComputerName     = $backupComputerName
            Account          = $backupAccount
        }

        Write-Log "iDrive: $idriveStatus | Last backup: $lastBackupDate | Result: $lastBackupInfo | Files: $backupFilesCount"

    } else {
        $Results.iDriveBackup = @{
            Status    = "NOT INSTALLED"
            Installed = $false
        }
        Write-Log "iDrive client not found on this system"
    }

} catch {
    Write-Log "iDrive check failed: $_" "ERROR"
    $Results.iDriveBackup = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "iDrive: $_"
}


# ============================================================================
# MODULE 5B: MALWAREBYTES ENDPOINT PROTECTION
# ============================================================================
Write-Log "Checking Malwarebytes status..."

try {
    $mbInstalled = $false
    $mbVersion = "N/A"
    $mbProductType = "N/A"
    $mbServiceRunning = $false
    $mbRealTimeProtection = "N/A"
    $mbLastScan = "N/A"
    $mbDefinitionsAge = "N/A"
    $mbLicenseStatus = "N/A"

    # Check for Malwarebytes service
    $mbServices = Get-Service -Name "MBAMService", "MBEndpointAgent" -ErrorAction SilentlyContinue
    $mbServiceRunning = ($mbServices | Where-Object { $_.Status -eq "Running" }).Count -gt 0

    # Check install paths
    $mbPaths = @(
        "$env:ProgramFiles\Malwarebytes\Anti-Malware\mbam.exe",
        "${env:ProgramFiles(x86)}\Malwarebytes\Anti-Malware\mbam.exe",
        "$env:ProgramFiles\Malwarebytes Endpoint Agent\MBEndpointAgent.exe",
        "$env:ProgramFiles\Malwarebytes\Malwarebytes\mbam.exe"
    )
    foreach ($mbPath in $mbPaths) {
        if (Test-Path $mbPath) {
            $mbInstalled = $true
            $mbFileInfo = Get-Item $mbPath -ErrorAction SilentlyContinue
            if ($mbFileInfo.VersionInfo.ProductVersion) {
                $mbVersion = $mbFileInfo.VersionInfo.ProductVersion
            }
            break
        }
    }

    # Also check via services
    if ($mbServices) { $mbInstalled = $true }

    if ($mbInstalled) {
        # Determine product type (Premium/Free/Endpoint)
        $mbEndpointSvc = Get-Service -Name "MBEndpointAgent" -ErrorAction SilentlyContinue
        if ($mbEndpointSvc) {
            $mbProductType = "Endpoint Protection"
        } else {
            # Check registry for license type
            $mbRegPaths = @(
                "HKLM:\SOFTWARE\Malwarebytes\Malwarebytes",
                "HKLM:\SOFTWARE\Malwarebytes' Anti-Malware"
            )
            foreach ($regPath in $mbRegPaths) {
                if (Test-Path $regPath) {
                    $mbReg = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($mbReg) {
                        if ($mbReg.premium -eq 1 -or $mbReg.IsPremium -eq 1) {
                            $mbProductType = "Premium"
                        } else {
                            $mbProductType = "Free"
                        }
                        break
                    }
                }
            }
        }

        # Check real-time protection via WMI (registered antivirus products)
        try {
            $avProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction Stop
            $mbAV = $avProducts | Where-Object { $_.displayName -match "Malwarebytes" }
            if ($mbAV) {
                # productState is a bitmask: bits 12-15 indicate real-time protection
                $state = $mbAV.productState
                $rtEnabled = (($state -shr 12) -band 0xF) -eq 1
                $mbRealTimeProtection = if ($rtEnabled) { "Enabled" } else { "Disabled" }
            }
        } catch { }

        # Check last scan and definitions from logs
        $mbLogsDir = "$env:ProgramData\Malwarebytes\MBAMService\logs"
        if (Test-Path $mbLogsDir) {
            $mbLogFile = Get-ChildItem -LiteralPath $mbLogsDir -Filter "mbamservice.log" -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($mbLogFile) {
                $mbLogTail = Get-Content -LiteralPath $mbLogFile.FullName -Tail 200 -ErrorAction SilentlyContinue | Out-String
                if ($mbLogTail -match "Scan completed") {
                    $mbLastScan = $mbLogFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                }
            }
        }

        # Check database/definitions age
        $mbDbPaths = @(
            "$env:ProgramData\Malwarebytes\MBAMService\data\mbam-rules.db",
            "$env:ProgramData\Malwarebytes\MBAMService\SignatureUpdates"
        )
        foreach ($dbPath in $mbDbPaths) {
            if (Test-Path $dbPath) {
                $dbItem = Get-Item $dbPath -ErrorAction SilentlyContinue
                if ($dbItem) {
                    $daysOld = [math]::Round(((Get-Date) - $dbItem.LastWriteTime).TotalDays, 1)
                    $mbDefinitionsAge = "$daysOld day(s) ago"
                    break
                }
            }
        }

        $mbStatus = "INSTALLED"
        if (-not $mbServiceRunning) {
            $mbStatus = "WARNING - Service not running"
        }

        $Results.Malwarebytes = @{
            Installed           = $true
            Status              = $mbStatus
            Version             = $mbVersion
            ProductType         = $mbProductType
            ServiceRunning      = $mbServiceRunning
            RealTimeProtection  = $mbRealTimeProtection
            LastScan            = $mbLastScan
            DefinitionsAge      = $mbDefinitionsAge
        }

        Write-Log "Malwarebytes: $mbProductType v$mbVersion | Service: $(if($mbServiceRunning){'Running'}else{'Stopped'}) | RT: $mbRealTimeProtection"
    } else {
        $Results.Malwarebytes = @{
            Installed = $false
            Status    = "NOT INSTALLED"
        }
        Write-Log "Malwarebytes not found on this system"
    }

} catch {
    Write-Log "Malwarebytes check failed: $_" "ERROR"
    $Results.Malwarebytes = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Malwarebytes: $_"
}


# ============================================================================
# MODULE 6: WINDOWS DEFENDER / ANTIVIRUS STATUS
# ============================================================================
Write-Log "Checking Windows Defender / Antivirus status..."

try {
    $defenderStatus = "Unknown"
    $rtProtection = "Unknown"
    $defDefs = "Unknown"
    $defsAge = "N/A"
    $lastScan = "Unknown"
    $lastScanType = "N/A"
    $threatCount = 0
    $recentThreats = @()

    # Get Defender status via Get-MpComputerStatus
    $mpStatus = Get-MpComputerStatus -ErrorAction Stop

    # Real-time protection
    $rtProtection = if ($mpStatus.RealTimeProtectionEnabled) { "Enabled" } else { "DISABLED" }

    # Definition status
    $defsDate = $mpStatus.AntivirusSignatureLastUpdated
    if ($defsDate) {
        $defDefs = $defsDate.ToString("yyyy-MM-dd HH:mm:ss")
        $defsAgeDays = [math]::Round(((Get-Date) - $defsDate).TotalDays, 1)
        $defsAge = "$defsAgeDays day(s) ago"
    }

    # Last scan info
    if ($mpStatus.FullScanEndTime -and $mpStatus.FullScanEndTime -gt [DateTime]::MinValue) {
        $lastScan = $mpStatus.FullScanEndTime.ToString("yyyy-MM-dd HH:mm:ss")
        $lastScanType = "Full Scan"
    }
    if ($mpStatus.QuickScanEndTime -and $mpStatus.QuickScanEndTime -gt [DateTime]::MinValue) {
        $quickDate = $mpStatus.QuickScanEndTime
        if ($lastScan -eq "Unknown" -or $quickDate -gt [DateTime]::Parse($lastScan)) {
            $lastScan = $quickDate.ToString("yyyy-MM-dd HH:mm:ss")
            $lastScanType = "Quick Scan"
        }
    }

    # Recent threats (last 30 days)
    try {
        $threats = Get-MpThreatDetection -ErrorAction SilentlyContinue |
            Where-Object { $_.InitialDetectionTime -gt (Get-Date).AddDays(-30) }
        if ($threats) {
            $threatCount = ($threats | Measure-Object).Count
            $recentThreats = $threats | Select-Object -First 5 | ForEach-Object {
                @{
                    Name = (Get-MpThreat -ThreatID $_.ThreatID -ErrorAction SilentlyContinue).ThreatName
                    Date = $_.InitialDetectionTime.ToString("yyyy-MM-dd HH:mm")
                    Action = $_.ActionSuccess
                }
            }
        }
    } catch {
        Write-Log "Could not retrieve threat history: $_" "WARN"
    }

    # Determine overall status
    if (-not $mpStatus.RealTimeProtectionEnabled) {
        # Check if a third-party AV (e.g. Malwarebytes) is providing protection instead
        $thirdPartyAV = $null
        try {
            $avProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction SilentlyContinue
            $thirdPartyAV = $avProducts | Where-Object { $_.displayName -notmatch "Windows Defender|Microsoft Defender" }
        } catch {}

        if ($thirdPartyAV) {
            $avNames = ($thirdPartyAV | ForEach-Object { $_.displayName }) -join ", "
            $defenderStatus = "PASS - Defender disabled ($avNames active)"
            $rtProtection = "Disabled - $avNames active"
        } else {
            $defenderStatus = "FAIL - Real-time protection DISABLED"
        }
    } elseif ($defsAgeDays -gt 7) {
        $defenderStatus = "WARNING - Definitions $defsAgeDays days old"
    } elseif ($threatCount -gt 0) {
        $defenderStatus = "WARNING - $threatCount threat(s) detected in last 30 days"
    } else {
        $defenderStatus = "PASS"
    }

    $Results.Defender = @{
        Status              = $defenderStatus
        RealTimeProtection  = $rtProtection
        DefinitionsUpdated  = $defDefs
        DefinitionsAge      = $defsAge
        LastScan            = $lastScan
        LastScanType        = $lastScanType
        ThreatsLast30Days   = $threatCount
        RecentThreats       = $recentThreats
        EngineVersion       = $mpStatus.AMEngineVersion
        ProductVersion      = $mpStatus.AMProductVersion
    }

    Write-Log "Defender: $defenderStatus | RT: $rtProtection | Defs: $defsAge | Threats (30d): $threatCount"

} catch {
    # Defender may not be available (e.g., third-party AV installed)
    Write-Log "Windows Defender not available or error: $_" "WARN"

    # Try to check if any antivirus is registered via WMI
    try {
        $avProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction SilentlyContinue
        if ($avProducts) {
            $avNames = ($avProducts | ForEach-Object { $_.displayName }) -join ", "
            $Results.Defender = @{
                Status             = "INFO - Third-party AV detected"
                RealTimeProtection = "See third-party AV"
                ThirdPartyAV       = $avNames
                ThreatsLast30Days  = 0
                RecentThreats      = @()
            }
            Write-Log "Third-party antivirus detected: $avNames"
        } else {
            $Results.Defender = @{
                Status             = "WARNING - No antivirus detected"
                RealTimeProtection = "Unknown"
                ThreatsLast30Days  = 0
                RecentThreats      = @()
            }
            Write-Log "No antivirus product detected" "WARN"
        }
    } catch {
        $Results.Defender = @{ Status = "ERROR"; Error = $_.ToString(); ThreatsLast30Days = 0; RecentThreats = @() }
        $Results.Errors += "Defender: $_"
    }
}


# ============================================================================
# MODULE 7: CRITICAL EVENT LOG ERRORS (Last 24 hours)
# ============================================================================
Write-Log "Checking Windows Event Logs for critical errors..."

try {
    $hoursBack = 24
    $cutoffTime = (Get-Date).AddHours(-$hoursBack)

    # ---- Known Event ID lookup table ----
    # Each entry: Source+EventID = @{ Category = "Routine"|"Actionable"; Note = "plain English explanation" }
    # Category "Routine" = safe to ignore, "Actionable" = needs attention
    $knownEvents = @{
        # DCOM / DistributedCOM permission errors - extremely common, safe to ignore
        "DistributedCOM:10016" = @{ Category = "Routine"; Note = "DCOM permission warning - cosmetic, safe to ignore. Windows components requesting access they already have." }
        "DCOM:10016"           = @{ Category = "Routine"; Note = "DCOM permission warning - cosmetic, safe to ignore." }
        # ESENT database engine warnings (used by Windows Search, Edge, etc.)
        "ESENT:455"            = @{ Category = "Routine"; Note = "ESENT database write warning - usually Windows Search or Edge temp database. Resolves on its own." }
        "ESENT:490"            = @{ Category = "Routine"; Note = "ESENT database recovery - Windows Search index self-repair. Normal operation." }
        "ESENT:489"            = @{ Category = "Routine"; Note = "ESENT database recovery - Windows Search index self-repair. Normal operation." }
        "ESENT:454"            = @{ Category = "Routine"; Note = "ESENT database operation - normal Search/Edge database maintenance." }
        "ESENT:642"            = @{ Category = "Routine"; Note = "ESENT database parameter change - normal Search/Edge operation." }
        # WMI errors
        "WMI:10"               = @{ Category = "Routine"; Note = "WMI filter error - usually a timing issue during startup. Safe to ignore unless repeated constantly." }
        "Microsoft-Windows-WMI:10"  = @{ Category = "Routine"; Note = "WMI filter error - timing issue during startup, safe to ignore." }
        # User Profile Service
        "Microsoft-Windows-User Profiles Service:1542" = @{ Category = "Routine"; Note = "User profile registry hive unloaded while still in use - usually harmless, happens during logoff." }
        # Windows Search
        "Microsoft-Windows-Search:3104" = @{ Category = "Routine"; Note = "Windows Search indexer failed to process an item - usually temp files that no longer exist." }
        "Microsoft-Windows-Search:3100" = @{ Category = "Routine"; Note = "Windows Search indexer error - usually self-resolving." }
        # Application Error / Hang
        "Application Error:1000" = @{ Category = "Actionable"; Note = "Application crash detected - check which program crashed. May need updating or reinstalling." }
        "Application Hang:1002"  = @{ Category = "Actionable"; Note = "Application stopped responding - check which program hung. May indicate insufficient RAM or failing disk." }
        # Windows Error Reporting
        "Windows Error Reporting:1001" = @{ Category = "Routine"; Note = "Windows Error Report submitted - this is the report about a crash, not the crash itself." }
        # Kernel / System critical
        "Microsoft-Windows-Kernel-Power:41" = @{ Category = "Actionable"; Note = "CRITICAL: Unexpected shutdown or power loss. Could indicate power supply issue, overheating, or system freeze." }
        "Microsoft-Windows-Kernel-Power:137" = @{ Category = "Routine"; Note = "Firmware performance counter issue - informational, safe to ignore." }
        "EventLog:6008"        = @{ Category = "Actionable"; Note = "Previous shutdown was unexpected - system crashed or lost power. Investigate if recurring." }
        # Disk errors
        "disk:11"              = @{ Category = "Actionable"; Note = "Disk I/O error detected - could indicate a failing hard drive. Monitor closely." }
        "disk:7"               = @{ Category = "Actionable"; Note = "Bad block detected on disk - potential drive failure. Run disk diagnostics." }
        "Ntfs:55"              = @{ Category = "Actionable"; Note = "NTFS file system corruption detected - run chkdsk. May indicate failing drive." }
        "Ntfs:137"             = @{ Category = "Actionable"; Note = "NTFS file system error on volume - investigate disk health." }
        # Windows Update
        "Microsoft-Windows-WindowsUpdateClient:20" = @{ Category = "Actionable"; Note = "Windows Update installation failed - may need manual intervention or troubleshooting." }
        "Microsoft-Windows-WindowsUpdateClient:25" = @{ Category = "Routine"; Note = "Windows Update download failed - will usually retry automatically." }
        # Service Control Manager
        "Service Control Manager:7000" = @{ Category = "Actionable"; Note = "A Windows service failed to start - check which service and whether it is needed." }
        "Service Control Manager:7001" = @{ Category = "Actionable"; Note = "A service depends on another service that failed - check the dependency chain." }
        "Service Control Manager:7009" = @{ Category = "Routine"; Note = "Service start timeout - service took too long to respond. Often happens after updates, usually resolves after reboot." }
        "Service Control Manager:7011" = @{ Category = "Routine"; Note = "Service timeout waiting for transaction response - usually a slow startup, not a real problem." }
        "Service Control Manager:7023" = @{ Category = "Actionable"; Note = "Service terminated with an error - check which service failed and why." }
        "Service Control Manager:7031" = @{ Category = "Actionable"; Note = "Service crashed and recovery action was taken - investigate if this keeps happening." }
        "Service Control Manager:7034" = @{ Category = "Actionable"; Note = "Service terminated unexpectedly - has happened multiple times, needs investigation." }
        # Schannel / TLS
        "Schannel:36887"       = @{ Category = "Routine"; Note = "TLS/SSL alert received from remote server - usually a website or service with certificate issues, not your PC." }
        "Schannel:36874"       = @{ Category = "Routine"; Note = "TLS connection attempt failed - remote server rejected connection. Usually a web server issue." }
        # VSS (Volume Shadow Copy)
        "VSS:12289"            = @{ Category = "Routine"; Note = "VSS snapshot issue - usually happens when backup software or System Restore runs during heavy disk activity." }
        "VSS:8194"             = @{ Category = "Routine"; Note = "VSS writer error - backup-related, usually resolves on retry." }
        # Print / Spooler
        "Microsoft-Windows-PrintService:372" = @{ Category = "Routine"; Note = "Print spooler error - usually a stale print job or driver issue. Restarting Print Spooler fixes it." }
        # Perflib / Performance counters
        "Perflib:1008"         = @{ Category = "Routine"; Note = "Performance counter library error - cosmetic, does not affect system performance." }
        "Perflib:1023"         = @{ Category = "Routine"; Note = "Performance counter library issue - informational only." }
        # Security-SPP (Licensing)
        "Security-SPP:16394"   = @{ Category = "Routine"; Note = "Software licensing notification - routine license validation check." }
        "Security-SPP:8198"    = @{ Category = "Routine"; Note = "License activation issue - check Windows activation status if this persists." }
        # .NET Runtime
        ".NET Runtime:1026"    = @{ Category = "Actionable"; Note = ".NET application crash - check which application and consider updating or reinstalling it." }
        # Task Scheduler
        "Microsoft-Windows-TaskScheduler:101" = @{ Category = "Routine"; Note = "Scheduled task failed to start - often a one-off timing issue." }
        # Certificate
        "Microsoft-Windows-CertificateServicesClient-CertEnroll:86" = @{ Category = "Routine"; Note = "Certificate enrollment issue - usually auto-resolves on domain-joined machines." }
        # DNS Client
        "Microsoft-Windows-DNS-Client:1014" = @{ Category = "Routine"; Note = "DNS name resolution timeout - usually a temporary network or DNS server issue." }
        # Audit / Security
        "Microsoft-Windows-Security-Auditing:4625" = @{ Category = "Actionable"; Note = "Failed login attempt - could indicate unauthorized access attempt if repeated." }
        # WHEA (Hardware Errors)
        "WHEA-Logger:17"       = @{ Category = "Actionable"; Note = "Hardware error detected by Windows - could be CPU, RAM, or bus error. Monitor closely." }
        "WHEA-Logger:18"       = @{ Category = "Actionable"; Note = "Fatal hardware error - serious hardware issue detected. Investigate immediately." }
        "WHEA-Logger:19"       = @{ Category = "Actionable"; Note = "Corrected hardware error - hardware detected and corrected an error. Monitor for increasing frequency." }
        # BugCheck (BSOD)
        "Microsoft-Windows-WER-SystemErrorReporting:1001" = @{ Category = "Actionable"; Note = "Blue Screen of Death (BSOD) crash dump saved - investigate the stop code for root cause." }
        "BugCheck:1001"        = @{ Category = "Actionable"; Note = "Blue Screen (BSOD) occurred - check crash dump for the specific error code." }
    }

    $allEvents = @()
    $actionableEvents = @()
    $routineEvents = @()

    # Check System and Application logs for Critical and Error events
    $logNames = @("System", "Application")

    foreach ($logName in $logNames) {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = $logName
            Level     = @(1, 2)  # 1 = Critical, 2 = Error
            StartTime = $cutoffTime
        } -MaxEvents 50 -ErrorAction SilentlyContinue

        foreach ($evt in $events) {
            $evtSource = $evt.ProviderName
            $evtId = $evt.Id
            $lookupKey = "${evtSource}:${evtId}"

            # Look up in known events table
            $knownInfo = $knownEvents[$lookupKey]
            $category = "Unknown"
            $note = "Not in known events database - review the message for context."

            if ($knownInfo) {
                $category = $knownInfo.Category
                $note = $knownInfo.Note
            } else {
                # Try partial source match (some sources have varying prefixes)
                $matchFound = $false
                foreach ($key in $knownEvents.Keys) {
                    $keyParts = $key -split ":"
                    if ($keyParts.Count -eq 2 -and $evtSource -match [regex]::Escape($keyParts[0]) -and $evtId -eq [int]$keyParts[1]) {
                        $knownInfo = $knownEvents[$key]
                        $category = $knownInfo.Category
                        $note = $knownInfo.Note
                        $matchFound = $true
                        break
                    }
                }
                # If still unknown, apply heuristics
                if (-not $matchFound) {
                    # Critical events are always actionable
                    if ($evt.Level -eq 1) {
                        $category = "Actionable"
                        $note = "Critical event from $evtSource - investigate further."
                    }
                }
            }

            $eventEntry = @{
                Log       = $logName
                Level     = if ($evt.Level -eq 1) { "Critical" } else { "Error" }
                Source    = $evtSource
                EventID   = $evtId
                Time      = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                Message   = if ($evt.Message.Length -gt 200) { $evt.Message.Substring(0, 200) + "..." } else { $evt.Message }
                Category  = $category
                Note      = $note
            }

            $allEvents += $eventEntry

            if ($category -eq "Actionable") {
                $actionableEvents += $eventEntry
            } else {
                $routineEvents += $eventEntry
            }
        }
    }

    # Count by severity and category
    $critCount = ($allEvents | Where-Object { $_.Level -eq "Critical" } | Measure-Object).Count
    $errCount = ($allEvents | Where-Object { $_.Level -eq "Error" } | Measure-Object).Count
    $totalEvents = $allEvents.Count
    $actionableCount = $actionableEvents.Count
    $routineCount = $routineEvents.Count

    # Determine status based on ACTIONABLE events only
    $eventStatus = "PASS"
    if ($actionableCount -gt 0) {
        $actionCrit = ($actionableEvents | Where-Object { $_.Level -eq "Critical" } | Measure-Object).Count
        if ($actionCrit -gt 0) {
            $eventStatus = "WARNING - $actionableCount actionable event(s) ($actionCrit critical)"
        } else {
            $eventStatus = "WARNING - $actionableCount event(s) need attention"
        }
    } elseif ($totalEvents -gt 0) {
        $eventStatus = "PASS - $routineCount routine event(s), none need attention"
    }

    $Results.EventLogErrors = @{
        Status            = $eventStatus
        CriticalCount     = $critCount
        ErrorCount        = $errCount
        TotalEvents       = $totalEvents
        ActionableCount   = $actionableCount
        RoutineCount      = $routineCount
        HoursChecked      = $hoursBack
        ActionableEvents  = $actionableEvents | Select-Object -First 20
        RoutineEvents     = $routineEvents | Select-Object -First 10
        AllEvents         = $allEvents | Select-Object -First 30
    }

    Write-Log "Event Logs: $totalEvents total ($actionableCount actionable, $routineCount routine) in last ${hoursBack}h"

} catch {
    Write-Log "Event Log check failed: $_" "ERROR"
    $Results.EventLogErrors = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Event Logs: $_"
}


# ============================================================================
# MODULE 8: PENDING REBOOT CHECK
# ============================================================================
Write-Log "Checking for pending reboot..."

try {
    $rebootRequired = $false
    $rebootReasons = @()

    # Check Windows Update reboot flag
    $wuReboot = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
    if ($wuReboot) {
        $rebootRequired = $true
        $rebootReasons += "Windows Update requires reboot"
    }

    # Check Component Based Servicing
    $cbsReboot = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
    if ($cbsReboot) {
        $rebootRequired = $true
        $rebootReasons += "Component servicing requires reboot"
    }

    # Check pending file rename operations
    $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
    if ($pfro.PendingFileRenameOperations) {
        $rebootRequired = $true
        $rebootReasons += "Pending file rename operations"
    }

    # Check if computer name change is pending
    $activeCompName = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -ErrorAction SilentlyContinue).ComputerName
    $pendingCompName = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -ErrorAction SilentlyContinue).ComputerName
    if ($activeCompName -and $pendingCompName -and $activeCompName -ne $pendingCompName) {
        $rebootRequired = $true
        $rebootReasons += "Computer name change pending"
    }

    # Check SCCM client reboot flag
    try {
        $sccmReboot = Invoke-WmiMethod -Namespace "ROOT\ccm\ClientSDK" -Class "CCM_ClientUtilities" -Name "DetermineIfRebootPending" -ErrorAction SilentlyContinue
        if ($sccmReboot -and $sccmReboot.RebootPending) {
            $rebootRequired = $true
            $rebootReasons += "SCCM client requires reboot"
        }
    } catch { }

    $rebootStatus = if ($rebootRequired) { "WARNING - Reboot pending" } else { "PASS" }
    $reasonText = if ($rebootReasons.Count -gt 0) { $rebootReasons -join "; " } else { "None" }

    # Include uptime context
    $uptimeDays = $Results.SystemInfo.UptimeDays
    $uptimeWarning = $false
    if ($uptimeDays -gt 30) {
        $uptimeWarning = $true
        if (-not $rebootRequired) {
            $rebootStatus = "WARNING - No reboot in $uptimeDays days"
        }
    }

    $Results.PendingReboot = @{
        Status         = $rebootStatus
        RebootRequired = $rebootRequired
        Reasons        = $reasonText
        UptimeDays     = $uptimeDays
        UptimeWarning  = $uptimeWarning
    }

    Write-Log "Reboot: $rebootStatus | Reasons: $reasonText | Uptime: $uptimeDays days"

} catch {
    Write-Log "Pending reboot check failed: $_" "ERROR"
    $Results.PendingReboot = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Pending Reboot: $_"
}


# ============================================================================
# MODULE 9: WINDOWS FIREWALL STATUS
# ============================================================================
Write-Log "Checking Windows Firewall status..."

try {
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $fwResults = @()
    $fwAllEnabled = $true

    foreach ($profile in $fwProfiles) {
        $enabled = $profile.Enabled
        if (-not $enabled) { $fwAllEnabled = $false }
        $fwResults += @{
            Profile = $profile.Name
            Enabled = $enabled
            DefaultInboundAction  = $profile.DefaultInboundAction.ToString()
            DefaultOutboundAction = $profile.DefaultOutboundAction.ToString()
        }
    }

    $fwStatus = if ($fwAllEnabled) { "PASS" } else { "WARNING - One or more profiles DISABLED" }

    $Results.Firewall = @{
        Status    = $fwStatus
        AllEnabled = $fwAllEnabled
        Profiles  = $fwResults
    }

    foreach ($p in $fwResults) {
        $enabledText = if ($p.Enabled) { "ON" } else { "OFF" }
        Write-Log "  Firewall $($p.Profile): $enabledText (In: $($p.DefaultInboundAction), Out: $($p.DefaultOutboundAction))"
    }

} catch {
    Write-Log "Firewall check failed: $_" "ERROR"
    $Results.Firewall = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Firewall: $_"
}


# ============================================================================
# MODULE 10: USER ACCOUNTS AUDIT
# ============================================================================
Write-Log "Auditing local user accounts..."

try {
    $localUsers = Get-LocalUser -ErrorAction Stop
    $localAdmins = @()
    $allAccounts = @()
    $flaggedAccounts = @()

    # Get members of the Administrators group
    $adminMembers = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue

    # Known/expected admin account patterns
    $expectedAdmins = @(
        "Administrator",
        "DefaultAccount",
        "WDAGUtilityAccount"
    )

    foreach ($user in $localUsers) {
        $isAdmin = $false
        if ($adminMembers) {
            $isAdmin = ($adminMembers | Where-Object { $_.Name -like "*\$($user.Name)" }) -ne $null
        }

        $accountInfo = @{
            Name         = $user.Name
            Enabled      = $user.Enabled
            IsAdmin      = $isAdmin
            LastLogon    = if ($user.LastLogon) { $user.LastLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            PasswordSet  = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
            Description  = $user.Description
            Flagged      = $false
            FlagReason   = ""
        }

        # Flag unexpected enabled admin accounts
        if ($isAdmin -and $user.Enabled) {
            $isExpected = $false
            foreach ($exp in $expectedAdmins) {
                if ($user.Name -eq $exp) { $isExpected = $true; break }
            }
            # Also treat the first user account as expected (likely the owner)
            # We flag only if the account name looks suspicious
            if (-not $isExpected) {
                # Check for suspicious patterns: generic names, remote access hints
                $suspiciousPatterns = @("admin", "test", "temp", "guest", "support", "remote", "helpdesk", "service")
                $isSuspicious = $false
                foreach ($pattern in $suspiciousPatterns) {
                    if ($user.Name -like "*$pattern*" -and $user.Name -ne "Administrator") {
                        $isSuspicious = $true
                        break
                    }
                }
                if ($isSuspicious) {
                    $accountInfo.Flagged = $true
                    $accountInfo.FlagReason = "Enabled admin account with suspicious name"
                    $flaggedAccounts += $accountInfo
                }
            }
        }

        # Flag enabled accounts with no password ever set
        if ($user.Enabled -and -not $user.PasswordLastSet -and $user.Name -ne "DefaultAccount" -and $user.Name -ne "WDAGUtilityAccount") {
            $accountInfo.Flagged = $true
            $accountInfo.FlagReason = "Enabled account with no password set"
            if (-not ($flaggedAccounts | Where-Object { $_.Name -eq $user.Name })) {
                $flaggedAccounts += $accountInfo
            }
        }

        if ($isAdmin -and $user.Enabled) {
            $localAdmins += $accountInfo
        }
        $allAccounts += $accountInfo
    }

    $enabledCount = ($localUsers | Where-Object { $_.Enabled } | Measure-Object).Count
    $adminCount = $localAdmins.Count
    $flaggedCount = $flaggedAccounts.Count

    $accountStatus = "PASS"
    if ($flaggedCount -gt 0) {
        $accountStatus = "WARNING - $flaggedCount account(s) flagged for review"
    }

    $Results.UserAccounts = @{
        Status          = $accountStatus
        TotalAccounts   = $localUsers.Count
        EnabledCount    = $enabledCount
        AdminCount      = $adminCount
        FlaggedCount    = $flaggedCount
        AdminAccounts   = $localAdmins
        FlaggedAccounts = $flaggedAccounts
        AllAccounts     = $allAccounts
    }

    Write-Log "User Accounts: $enabledCount enabled, $adminCount admins, $flaggedCount flagged"

} catch {
    Write-Log "User accounts audit failed: $_" "ERROR"
    $Results.UserAccounts = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "User Accounts: $_"
}


# ============================================================================
# MODULE 11: TEMP FILES REPORT
# ============================================================================
Write-Log "Checking temp file sizes..."

try {
    $tempLocations = @()

    # Windows Temp
    $winTemp = "$env:SystemRoot\Temp"
    if (Test-Path $winTemp) {
        $size = (Get-ChildItem -LiteralPath $winTemp -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $sizeMB = [math]::Round($size / 1MB, 1)
        $fileCount = (Get-ChildItem -LiteralPath $winTemp -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
        $tempLocations += @{ Location = "Windows Temp"; Path = $winTemp; SizeMB = $sizeMB; FileCount = $fileCount }
    }

    # User Temp folders (all users)
    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch "^(Public|Default|Default User|All Users)$" }
    foreach ($profile in $userProfiles) {
        $userTemp = Join-Path $profile.FullName "AppData\Local\Temp"
        if (Test-Path $userTemp) {
            $size = (Get-ChildItem -LiteralPath $userTemp -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            $fileCount = (Get-ChildItem -LiteralPath $userTemp -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
            $tempLocations += @{ Location = "User Temp ($($profile.Name))"; Path = $userTemp; SizeMB = $sizeMB; FileCount = $fileCount }
        }
    }

    # Windows Update cache
    $wuCache = "$env:SystemRoot\SoftwareDistribution\Download"
    if (Test-Path $wuCache) {
        $size = (Get-ChildItem -LiteralPath $wuCache -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $sizeMB = [math]::Round($size / 1MB, 1)
        $fileCount = (Get-ChildItem -LiteralPath $wuCache -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
        $tempLocations += @{ Location = "Windows Update Cache"; Path = $wuCache; SizeMB = $sizeMB; FileCount = $fileCount }
    }

    # Browser caches (Chrome, Edge, Firefox)
    foreach ($profile in $userProfiles) {
        # Chrome
        $chromePath = Join-Path $profile.FullName "AppData\Local\Google\Chrome\User Data\Default\Cache"
        if (Test-Path $chromePath) {
            $size = (Get-ChildItem -LiteralPath $chromePath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            $tempLocations += @{ Location = "Chrome Cache ($($profile.Name))"; Path = $chromePath; SizeMB = $sizeMB; FileCount = 0 }
        }
        # Edge
        $edgePath = Join-Path $profile.FullName "AppData\Local\Microsoft\Edge\User Data\Default\Cache"
        if (Test-Path $edgePath) {
            $size = (Get-ChildItem -LiteralPath $edgePath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            $tempLocations += @{ Location = "Edge Cache ($($profile.Name))"; Path = $edgePath; SizeMB = $sizeMB; FileCount = 0 }
        }
        # Firefox
        $ffProfiles = Join-Path $profile.FullName "AppData\Local\Mozilla\Firefox\Profiles"
        if (Test-Path $ffProfiles) {
            $ffDirs = Get-ChildItem -LiteralPath $ffProfiles -Directory -ErrorAction SilentlyContinue
            foreach ($ffDir in $ffDirs) {
                $ffCache = Join-Path $ffDir.FullName "cache2"
                if (Test-Path $ffCache) {
                    $size = (Get-ChildItem -LiteralPath $ffCache -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    $sizeMB = [math]::Round($size / 1MB, 1)
                    $tempLocations += @{ Location = "Firefox Cache ($($profile.Name))"; Path = $ffCache; SizeMB = $sizeMB; FileCount = 0 }
                }
            }
        }
    }

    $totalSizeMB = 0
    foreach ($loc in $tempLocations) { $totalSizeMB += $loc.SizeMB }
    $totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)

    $tempStatus = "PASS"
    if ($totalSizeMB -gt 5120) {
        $tempStatus = "WARNING - ${totalSizeGB}GB of temp/cache files"
    } elseif ($totalSizeMB -gt 2048) {
        $tempStatus = "PASS - ${totalSizeGB}GB of temp/cache files"
    }

    $Results.TempFiles = @{
        Status        = $tempStatus
        TotalSizeMB   = $totalSizeMB
        TotalSizeGB   = $totalSizeGB
        Locations     = $tempLocations
    }

    Write-Log "Temp Files: ${totalSizeGB}GB total across $($tempLocations.Count) locations"

} catch {
    Write-Log "Temp files check failed: $_" "ERROR"
    $Results.TempFiles = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Temp Files: $_"
}


# ============================================================================
# MODULE 12: STARTUP PROGRAMS AUDIT
# ============================================================================
Write-Log "Auditing startup programs..."

try {
    $startupItems = @()

    # Registry: HKLM Run
    $regPaths = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; Scope = "All Users (HKLM)" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"; Scope = "All Users RunOnce" },
        @{ Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; Scope = "Current User (HKCU)" },
        @{ Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"; Scope = "Current User RunOnce" },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"; Scope = "All Users (32-bit)" }
    )

    foreach ($reg in $regPaths) {
        if (Test-Path $reg.Path) {
            $entries = Get-ItemProperty -Path $reg.Path -ErrorAction SilentlyContinue
            $entries.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
                $startupItems += @{
                    Name     = $_.Name
                    Command  = $_.Value
                    Source   = $reg.Scope
                    Type     = "Registry"
                }
            }
        }
    }

    # Startup folders
    $startupFolders = @(
        [Environment]::GetFolderPath("Startup"),
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and $_.Trim() -ne '' }

    foreach ($folder in $startupFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -File -ErrorAction SilentlyContinue | ForEach-Object {
                $startupItems += @{
                    Name    = $_.BaseName
                    Command = $_.FullName
                    Source  = "Startup Folder"
                    Type    = "Shortcut/File"
                }
            }
        }
    }

    # Scheduled tasks that run at logon
    $logonTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.Triggers | Where-Object { $_ -is [CimInstance] -and $_.CimClass.CimClassName -eq "MSFT_TaskLogonTrigger" } } |
        Select-Object TaskName, TaskPath, State

    foreach ($task in $logonTasks) {
        $startupItems += @{
            Name    = $task.TaskName
            Command = $task.TaskPath
            Source  = "Scheduled Task (Logon)"
            Type    = "ScheduledTask"
        }
    }

    # Filter out entries with empty/null commands
    $startupItems = @($startupItems | Where-Object { $_.Command -and "$($_.Command)".Trim() -ne '' })

    # Now flag suspicious items
    $flaggedItems = @()
    $safeItems = @()

    foreach ($item in $startupItems) {
        $isSafe = $false
        foreach ($safe in $KnownSafePrograms) {
            if ($item.Name -match [regex]::Escape($safe) -or $item.Command -match [regex]::Escape($safe)) {
                $isSafe = $true
                break
            }
        }

        $isSuspicious = $false
        $suspiciousReasons = @()

        if (-not $isSafe) {
            # Check against suspicious keywords
            foreach ($keyword in $SuspiciousKeywords) {
                if ($item.Name -match $keyword -or $item.Command -match $keyword) {
                    $isSuspicious = $true
                    $suspiciousReasons += "Matches keyword: $keyword"
                }
            }

            # Check for executables in unusual locations
            if ($item.Command -match "\\Temp\\|\\AppData\\Local\\Temp|\\Downloads\\") {
                $isSuspicious = $true
                $suspiciousReasons += "Runs from temp/downloads folder"
            }

            # Check for unsigned or unknown executables
            $exePath = $null
            if ($item.Command -and ($item.Command -match '"([^"]+\.exe)"' -or $item.Command -match '([^\s]+\.exe)')) {
                $exePath = $matches[1]
            }

            if ($exePath -and $exePath.Trim() -ne '' -and (Test-Path $exePath)) {
                $sig = Get-AuthenticodeSignature $exePath -ErrorAction SilentlyContinue
                if ($sig -and $sig.Status -ne "Valid") {
                    $isSuspicious = $true
                    $suspiciousReasons += "Not digitally signed"
                }
            }
        }

        $item.IsSuspicious = $isSuspicious
        $item.Reasons = $suspiciousReasons -join "; "
        $item.IsSafe = $isSafe

        if ($isSuspicious) {
            $flaggedItems += $item
        } else {
            $safeItems += $item
        }
    }

    $Results.StartupPrograms = @{
        Status       = if ($flaggedItems.Count -gt 0) { "WARNING  - $($flaggedItems.Count) suspicious item(s)" } else { "PASS" }
        TotalCount   = $startupItems.Count
        FlaggedCount = $flaggedItems.Count
        FlaggedItems = $flaggedItems
        AllItems     = $startupItems
    }

    Write-Log "Startup audit: $($startupItems.Count) items found, $($flaggedItems.Count) flagged as suspicious"
    foreach ($flagged in $flaggedItems) {
        Write-Log "  FLAGGED: $($flagged.Name)  - $($flagged.Reasons)" "WARN"
    }

} catch {
    Write-Log "Startup audit failed: $_" "ERROR"
    $Results.StartupPrograms = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Startup Audit: $_"
}


# ============================================================================
# MODULE 13: SCHEDULED TASKS AUDIT
# ============================================================================
Write-Log "Auditing scheduled tasks..."

try {
    # Known trusted publishers / task path patterns
    $trustedPublishers = @(
        "Microsoft", "Google", "Adobe", "Intel", "NVIDIA", "Realtek",
        "IDrive", "TeamViewer", "Dropbox", "Mozilla", "Brave",
        "PCMasterclass", "Apple", "Logitech", "Dell", "HP", "Lenovo",
        "ASUS", "Acer", "Samsung", "Oracle", "Zoom", "Citrix"
    )
    $trustedPathPrefixes = @("\Microsoft\")

    # Known-safe commands that should never be flagged even with unknown publisher
    $knownSafeCommands = @(
        "MicrosoftEdgeUpdate\.exe",
        "PCMasterclass-Maintenance\.ps1",
        "Apple Software Update\\SoftwareUpdate\.exe",
        "Lenovo\.Modern\.ImController\.exe"
    )

    $allTasks = Get-ScheduledTask -ErrorAction Stop
    $systemTasks = $allTasks | Where-Object {
        $isTrustedPath = $false
        foreach ($prefix in $trustedPathPrefixes) {
            if ($_.TaskPath -like "$prefix*") { $isTrustedPath = $true; break }
        }
        $isTrustedPath
    }
    $thirdPartyTasks = $allTasks | Where-Object {
        $isTrustedPath = $false
        foreach ($prefix in $trustedPathPrefixes) {
            if ($_.TaskPath -like "$prefix*") { $isTrustedPath = $true; break }
        }
        -not $isTrustedPath
    }

    $flaggedTasks = @()
    $thirdPartyDetails = @()

    foreach ($task in $thirdPartyTasks) {
        $taskInfo = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue
        $actions = $task.Actions | ForEach-Object {
            $cmd = $_.Execute
            if ($_.Arguments) { $cmd += " $($_.Arguments)" }
            $cmd
        }
        $command = ($actions -join "; ").Substring(0, [math]::Min(($actions -join "; ").Length, 200))
        $publisher = if ($task.Author) { $task.Author } else { "Unknown" }
        $nextRun = if ($taskInfo -and $taskInfo.NextRunTime -and $taskInfo.NextRunTime.Year -gt 1999) {
            $taskInfo.NextRunTime.ToString("yyyy-MM-dd HH:mm")
        } else { "N/A" }

        $detail = @{
            Name      = $task.TaskName
            Path      = $task.TaskPath
            Publisher = $publisher
            State     = $task.State.ToString()
            Command   = $command
            NextRun   = $nextRun
            Flagged   = $false
            FlagReason = ""
        }

        # Flag suspicious tasks
        $reasons = @()
        $isTrusted = $false
        foreach ($tp in $trustedPublishers) {
            if ($publisher -like "*$tp*") { $isTrusted = $true; break }
        }
        # Check if the command matches a known-safe pattern
        if (-not $isTrusted) {
            foreach ($safePattern in $knownSafeCommands) {
                if ($command -match $safePattern) { $isTrusted = $true; break }
            }
        }
        if (-not $isTrusted -and $publisher -eq "Unknown") {
            $reasons += "Unknown publisher"
        }
        if ($command -match "\\Temp\\|\\tmp\\|%temp%") {
            $reasons += "Runs from Temp folder"
        }
        if ($command -match "-enc\s|-encodedcommand\s|-e\s") {
            $reasons += "Encoded PowerShell command"
        }
        if ($command -match "cmd\.exe.*\/c.*http|powershell.*http|wget|curl") {
            $reasons += "Downloads from internet"
        }

        if ($reasons.Count -gt 0) {
            $detail.Flagged = $true
            $detail.FlagReason = $reasons -join ", "
            $flaggedTasks += $detail
        }

        $thirdPartyDetails += $detail
    }

    $taskStatus = "PASS"
    if ($flaggedTasks.Count -gt 0) {
        $taskStatus = "WARNING - $($flaggedTasks.Count) suspicious task(s)"
    }

    $Results.ScheduledTasks = @{
        Status           = $taskStatus
        TotalTasks       = $allTasks.Count
        SystemTasks      = $systemTasks.Count
        ThirdPartyTasks  = $thirdPartyTasks.Count
        FlaggedCount     = $flaggedTasks.Count
        FlaggedTasks     = $flaggedTasks
        ThirdPartyDetails = $thirdPartyDetails
    }

    Write-Log "Scheduled Tasks: $($allTasks.Count) total ($($systemTasks.Count) system, $($thirdPartyTasks.Count) third-party, $($flaggedTasks.Count) flagged)"

} catch {
    Write-Log "Scheduled tasks audit failed: $_" "ERROR"
    $Results.ScheduledTasks = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Scheduled Tasks: $_"
}


# ============================================================================
# MODULE 14: BROWSER EXTENSION AUDIT
# ============================================================================
Write-Log "Auditing browser extensions..."

# Helper: resolve __MSG_ extension names from _locales
# Trusted browser extensions that should never be flagged (by ID or name pattern)
$TrustedExtensionIDs = @(
    "cjpalhdlnbpafiamejdnhcphjbkeiagm",  # uBlock Origin
    "ddkjiahejlhfcafbddmgiahcphecmpfh",  # uBlock Origin Lite
    "efaidnbmnnnibpcajpcglclefindmkaj",   # Adobe Acrobat
    "ghbmnnjooekpmoecnnnilnnbdlolhkhi",   # Google Docs Offline
    "nmmhkkegccagdldgiimedpiccmgmieda",   # Google Wallet
    "pkedcjkdefgpdelpbcmbmeomcjbeemfm"    # Chrome Remote Desktop
)

# Well-known extension IDs whose names often fail to resolve from locale files
$KnownExtensionNames = @{
    "efaidnbmnnnibpcajpcglclefindmkaj" = "Adobe Acrobat"
    "ghbmnnjooekpmoecnnnilnnbdlolhkhi" = "Google Docs Offline"
    "nmmhkkegccagdldgiimedpiccmgmieda" = "Google Wallet (Chrome Web Store Payments)"
    "aapocclcgogkmnckokdopfmhonfmgoek" = "Google Slides"
    "aohghmighlieiainnegkcijnfilokake" = "Google Docs"
    "felcaaldnbdncclmgdcncolpebgiejap" = "Google Calendar"
    "blpcfgokakmgnkcojhhkbfbldkacnbeo" = "YouTube"
    "gomekmidlodglbbmalcneegieacbdmki" = "Avast Online Security & Privacy"
    "pkedcjkdefgpdelpbcmbmeomcjbeemfm" = "Chrome Remote Desktop"
    "cjpalhdlnbpafiamejdnhcphjbkeiagm" = "uBlock Origin"
}

function Resolve-ExtensionName {
    param([string]$ManifestName, [string]$ExtVersionDir, [string]$FallbackName)

    # Check known extension IDs first
    if ($KnownExtensionNames.ContainsKey($FallbackName)) {
        return $KnownExtensionNames[$FallbackName]
    }

    if (-not $ManifestName -or $ManifestName -match "^__MSG_") {
        # Extract the message key (e.g. __MSG_appName__ -> appName)
        $msgKey = $ManifestName -replace '^__MSG_', '' -replace '__$', ''

        # Try common locale folders in order
        $localeDirs = @("en", "en_US", "en_GB")
        foreach ($locale in $localeDirs) {
            $messagesPath = Join-Path $ExtVersionDir "_locales\$locale\messages.json"
            if (Test-Path $messagesPath) {
                try {
                    $messages = Get-Content -LiteralPath $messagesPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    # Try the exact key first, then common fallbacks
                    foreach ($key in @($msgKey, "appName", "extName", "extensionName", "app_name", "name")) {
                        if ($messages.$key -and $messages.$key.message) {
                            return $messages.$key.message
                        }
                    }
                } catch { }
            }
        }
        return $FallbackName
    }
    return $ManifestName
}

try {
    $browserExtensions = @()
    $flaggedExtensions = @()
    $browsersFound = @()

    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch "^(Public|Default|Default User|All Users)$" }

    foreach ($profile in $userProfiles) {
        # Chrome extensions
        $chromeExtDir = Join-Path $profile.FullName "AppData\Local\Google\Chrome\User Data\Default\Extensions"
        if (Test-Path $chromeExtDir) {
            if ("Chrome" -notin $browsersFound) { $browsersFound += "Chrome" }
            $extFolders = Get-ChildItem -LiteralPath $chromeExtDir -Directory -ErrorAction SilentlyContinue
            foreach ($extFolder in $extFolders) {
                # Find the latest version subfolder with a manifest
                $versionDirs = Get-ChildItem -LiteralPath $extFolder.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
                foreach ($verDir in $versionDirs) {
                    $manifestPath = Join-Path $verDir.FullName "manifest.json"
                    if (Test-Path $manifestPath) {
                        try {
                            $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
                            $extName = Resolve-ExtensionName -ManifestName $manifest.name -ExtVersionDir $verDir.FullName -FallbackName $extFolder.Name
                            $extVersion = if ($manifest.version) { $manifest.version } else { "Unknown" }
                            $permissions = @()
                            if ($manifest.permissions) { $permissions += $manifest.permissions }
                            if ($manifest.host_permissions) { $permissions += $manifest.host_permissions }

                            $ext = @{
                                Browser     = "Chrome"
                                User        = $profile.Name
                                Name        = $extName
                                Version     = $extVersion
                                ID          = $extFolder.Name
                                Permissions = ($permissions -join ", ")
                                Flagged     = $false
                                FlagReason  = ""
                            }

                            # Flag suspicious extensions (skip trusted ones)
                            $reasons = @()
                            if ($extFolder.Name -notin $TrustedExtensionIDs) {
                                $hasAllUrls = ($permissions -join " ") -match "<all_urls>|\*://\*/\*"
                                $hasHistory = ($permissions -join " ") -match "history|webNavigation"
                                $hasTabs = ($permissions -join " ") -match "\btabs\b"
                                if ($hasAllUrls -and $hasHistory) {
                                    $reasons += "Access to all sites + browsing history"
                                }
                                if ($extName -match "coupon|shop|deal|discount|cashback|honey" -and $hasAllUrls) {
                                    $reasons += "Shopping extension with broad permissions"
                                }
                            }

                            if ($reasons.Count -gt 0) {
                                $ext.Flagged = $true
                                $ext.FlagReason = $reasons -join "; "
                                $flaggedExtensions += $ext
                            }

                            $browserExtensions += $ext
                        } catch {
                            # Skip unreadable manifests
                        }
                        break  # Only read the first (latest) version
                    }
                }
            }
        }

        # Edge extensions
        $edgeExtDir = Join-Path $profile.FullName "AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
        if (Test-Path $edgeExtDir) {
            if ("Edge" -notin $browsersFound) { $browsersFound += "Edge" }
            $extFolders = Get-ChildItem -LiteralPath $edgeExtDir -Directory -ErrorAction SilentlyContinue
            foreach ($extFolder in $extFolders) {
                $versionDirs = Get-ChildItem -LiteralPath $extFolder.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
                foreach ($verDir in $versionDirs) {
                    $manifestPath = Join-Path $verDir.FullName "manifest.json"
                    if (Test-Path $manifestPath) {
                        try {
                            $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
                            $extName = Resolve-ExtensionName -ManifestName $manifest.name -ExtVersionDir $verDir.FullName -FallbackName $extFolder.Name
                            $extVersion = if ($manifest.version) { $manifest.version } else { "Unknown" }
                            $permissions = @()
                            if ($manifest.permissions) { $permissions += $manifest.permissions }
                            if ($manifest.host_permissions) { $permissions += $manifest.host_permissions }

                            $ext = @{
                                Browser     = "Edge"
                                User        = $profile.Name
                                Name        = $extName
                                Version     = $extVersion
                                ID          = $extFolder.Name
                                Permissions = ($permissions -join ", ")
                                Flagged     = $false
                                FlagReason  = ""
                            }

                            $reasons = @()
                            if ($extFolder.Name -notin $TrustedExtensionIDs) {
                                $hasAllUrls = ($permissions -join " ") -match "<all_urls>|\*://\*/\*"
                                $hasHistory = ($permissions -join " ") -match "history|webNavigation"
                                if ($hasAllUrls -and $hasHistory) {
                                    $reasons += "Access to all sites + browsing history"
                                }
                                if ($extName -match "coupon|shop|deal|discount|cashback" -and $hasAllUrls) {
                                    $reasons += "Shopping extension with broad permissions"
                                }
                            }

                            if ($reasons.Count -gt 0) {
                                $ext.Flagged = $true
                                $ext.FlagReason = $reasons -join "; "
                                $flaggedExtensions += $ext
                            }

                            $browserExtensions += $ext
                        } catch { }
                        break
                    }
                }
            }
        }

        # Firefox extensions
        $ffProfileDir = Join-Path $profile.FullName "AppData\Roaming\Mozilla\Firefox\Profiles"
        if (Test-Path $ffProfileDir) {
            if ("Firefox" -notin $browsersFound) { $browsersFound += "Firefox" }
            $ffProfiles = Get-ChildItem -LiteralPath $ffProfileDir -Directory -ErrorAction SilentlyContinue
            foreach ($ffProf in $ffProfiles) {
                $extensionsJson = Join-Path $ffProf.FullName "extensions.json"
                if (Test-Path $extensionsJson) {
                    try {
                        $ffData = Get-Content -LiteralPath $extensionsJson -Raw -ErrorAction Stop | ConvertFrom-Json
                        foreach ($addon in $ffData.addons) {
                            if ($addon.type -eq "extension" -and $addon.location -ne "app-system-defaults") {
                                $ext = @{
                                    Browser     = "Firefox"
                                    User        = $profile.Name
                                    Name        = if ($addon.defaultLocale.name) { $addon.defaultLocale.name } else { $addon.id }
                                    Version     = if ($addon.version) { $addon.version } else { "Unknown" }
                                    ID          = $addon.id
                                    Permissions = ""
                                    Flagged     = $false
                                    FlagReason  = ""
                                }
                                if (-not $addon.active) {
                                    $ext.FlagReason = "Disabled extension"
                                }
                                $browserExtensions += $ext
                            }
                        }
                    } catch { }
                }
            }
        }
    }

    $extStatus = "PASS"
    if ($flaggedExtensions.Count -gt 0) {
        $extStatus = "WARNING - $($flaggedExtensions.Count) extension(s) flagged for review"
    }

    $Results.BrowserExtensions = @{
        Status           = $extStatus
        BrowsersFound    = $browsersFound -join ", "
        TotalExtensions  = $browserExtensions.Count
        FlaggedCount     = $flaggedExtensions.Count
        FlaggedExtensions = $flaggedExtensions
        AllExtensions    = $browserExtensions
    }

    Write-Log "Browser Extensions: $($browserExtensions.Count) total across $($browsersFound -join ', ') ($($flaggedExtensions.Count) flagged)"

} catch {
    Write-Log "Browser extension audit failed: $_" "ERROR"
    $Results.BrowserExtensions = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Browser Extensions: $_"
}


# ============================================================================
# MODULE 15: SERVICE STATUS MONITOR
# ============================================================================
Write-Log "Checking key service statuses..."

try {
    # Key services to monitor: ServiceName => Friendly description
    $keyServices = [ordered]@{
        "wuauserv"    = "Windows Update"
        "WinDefend"   = "Windows Defender"
        "mpssvc"      = "Windows Firewall"
        "Dhcp"        = "DHCP Client"
        "Dnscache"    = "DNS Client"
        "WSearch"     = "Windows Search"
        "Schedule"    = "Task Scheduler"
        "EventLog"    = "Event Log"
        "Spooler"     = "Print Spooler"
        "Winmgmt"     = "WMI (Management)"
        "BITS"        = "Background Intelligent Transfer"
        "CryptSvc"    = "Cryptographic Services"
        "W32Time"     = "Windows Time"
    }

    # Services that are OK to be stopped/disabled
    $acceptablyStopped = @("TermService", "RemoteRegistry", "Fax", "TapiSrv")

    $serviceResults = @()
    $stoppedUnexpected = @()

    foreach ($svcName in $keyServices.Keys) {
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if ($svc) {
            $svcDetail = @{
                Name        = $svcName
                DisplayName = $keyServices[$svcName]
                Status      = $svc.Status.ToString()
                StartType   = $svc.StartType.ToString()
                Flagged     = $false
                FlagReason  = ""
            }

            if ($svc.Status -ne "Running" -and $svcName -notin $acceptablyStopped) {
                if ($svc.StartType -eq "Automatic" -or $svc.StartType -eq "Manual") {
                    $svcDetail.Flagged = $true
                    $svcDetail.FlagReason = "Set to $($svc.StartType) but not running"
                    $stoppedUnexpected += $svcDetail
                }
            }

            $serviceResults += $svcDetail
        }
    }

    # Also check common third-party services
    $thirdPartyServices = @(
        "IDriveService", "IDrive*",
        "TeamViewer", "tvnserver",
        "DropboxUpdate", "dbupdate",
        "gupdate", "gupdatem",
        "AdobeARMservice",
        "NVDisplay.ContainerLocalSystem"
    )

    foreach ($svcPattern in $thirdPartyServices) {
        $services = Get-Service -Name $svcPattern -ErrorAction SilentlyContinue
        foreach ($svc in $services) {
            # Avoid duplicates
            if ($serviceResults | Where-Object { $_.Name -eq $svc.Name }) { continue }
            $svcDetail = @{
                Name        = $svc.Name
                DisplayName = $svc.DisplayName
                Status      = $svc.Status.ToString()
                StartType   = $svc.StartType.ToString()
                Flagged     = $false
                FlagReason  = ""
            }

            if ($svc.Status -ne "Running" -and $svc.StartType -eq "Automatic") {
                $svcDetail.Flagged = $true
                $svcDetail.FlagReason = "Set to Automatic but not running"
                $stoppedUnexpected += $svcDetail
            }

            $serviceResults += $svcDetail
        }
    }

    $svcStatus = "PASS"
    if ($stoppedUnexpected.Count -gt 0) {
        $svcStatus = "WARNING - $($stoppedUnexpected.Count) service(s) unexpectedly stopped"
    }

    $Results.ServiceStatus = @{
        Status              = $svcStatus
        TotalChecked        = $serviceResults.Count
        Running             = ($serviceResults | Where-Object { $_.Status -eq "Running" }).Count
        StoppedUnexpected   = $stoppedUnexpected.Count
        FlaggedServices     = $stoppedUnexpected
        AllServices         = $serviceResults
    }

    Write-Log "Services: $($serviceResults.Count) checked, $(($serviceResults | Where-Object { $_.Status -eq 'Running' }).Count) running, $($stoppedUnexpected.Count) unexpectedly stopped"

} catch {
    Write-Log "Service status check failed: $_" "ERROR"
    $Results.ServiceStatus = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Service Status: $_"
}


# ============================================================================
# MODULE 16: NETWORK CONFIGURATION CHECK
# ============================================================================
Write-Log "Running network configuration check..."

try {
    $netIssues = @()
    $netAdapters = @()

    # Get detailed adapter info including link speed
    $physicalAdapters = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.NetEnabled -eq $true -and $_.PhysicalAdapter -eq $true }
    $adapterConfigs = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

    foreach ($config in $adapterConfigs) {
        $physAdapter = $physicalAdapters | Where-Object { $_.Index -eq $config.Index }
        $adapterName = if ($physAdapter) { $physAdapter.Name } else { $config.Description }
        $linkSpeedMbps = if ($physAdapter -and $physAdapter.Speed) { [math]::Round($physAdapter.Speed / 1000000, 0) } else { "Unknown" }

        $ipv4Addrs = @($config.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
        $gateway = if ($config.DefaultIPGateway) { ($config.DefaultIPGateway | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ", " } else { "None" }
        $dns = if ($config.DNSServerSearchOrder) { $config.DNSServerSearchOrder -join ", " } else { "None" }
        $subnetMasks = @($config.IPSubnet | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })

        # DHCP lease info
        $dhcpServer = if ($config.DHCPEnabled -and $config.DHCPServer) { $config.DHCPServer } else { "N/A" }
        $dhcpLeaseObtained = if ($config.DHCPEnabled -and $config.DHCPLeaseObtained) { $config.DHCPLeaseObtained.ToString("yyyy-MM-dd HH:mm") } else { "N/A" }
        $dhcpLeaseExpires = if ($config.DHCPEnabled -and $config.DHCPLeaseExpires) { $config.DHCPLeaseExpires.ToString("yyyy-MM-dd HH:mm") } else { "N/A" }

        $adapterDetail = @{
            Name            = $adapterName
            IPv4            = $ipv4Addrs -join ", "
            SubnetMask      = $subnetMasks -join ", "
            Gateway         = $gateway
            DNS             = $dns
            MAC             = $config.MACAddress
            DHCP            = $config.DHCPEnabled
            DHCPServer      = $dhcpServer
            LeaseObtained   = $dhcpLeaseObtained
            LeaseExpires    = $dhcpLeaseExpires
            LinkSpeedMbps   = $linkSpeedMbps
        }
        $netAdapters += $adapterDetail

        # Issue detection: APIPA address (169.254.x.x)
        foreach ($ip in $ipv4Addrs) {
            if ($ip -match '^169\.254\.') {
                $netIssues += @{ Adapter = $adapterName; Issue = "APIPA address detected ($ip) - DHCP may be failing"; Severity = "ERROR" }
            }
        }

        # Issue detection: no gateway
        if ($gateway -eq "None" -or $gateway -eq "") {
            $netIssues += @{ Adapter = $adapterName; Issue = "No default gateway configured - may have no internet access"; Severity = "WARNING" }
        }

        # Issue detection: no DNS
        if ($dns -eq "None" -or $dns -eq "") {
            $netIssues += @{ Adapter = $adapterName; Issue = "No DNS servers configured - name resolution will fail"; Severity = "ERROR" }
        }

        # Issue detection: slow link speed (less than 100 Mbps for wired)
        if ($linkSpeedMbps -ne "Unknown" -and $linkSpeedMbps -lt 100 -and $adapterName -notmatch "Wi-Fi|Wireless|802\.11") {
            $netIssues += @{ Adapter = $adapterName; Issue = "Wired link speed is only $($linkSpeedMbps) Mbps - check cable or NIC"; Severity = "WARNING" }
        }
    }

    # Issue detection: multiple adapters on same subnet
    $subnets = @{}
    foreach ($adapter in $netAdapters) {
        $ips = $adapter.IPv4 -split ", "
        $masks = $adapter.SubnetMask -split ", "
        for ($i = 0; $i -lt $ips.Count; $i++) {
            if ($ips[$i] -and $masks[$i]) {
                $ipBytes = [System.Net.IPAddress]::Parse($ips[$i]).GetAddressBytes()
                $maskBytes = [System.Net.IPAddress]::Parse($masks[$i]).GetAddressBytes()
                $netBytes = @()
                for ($j = 0; $j -lt 4; $j++) { $netBytes += ($ipBytes[$j] -band $maskBytes[$j]) }
                $networkId = ($netBytes -join ".")
                if ($subnets.ContainsKey($networkId)) {
                    $netIssues += @{ Adapter = $adapter.Name; Issue = "Shares subnet $networkId with $($subnets[$networkId]) - may cause routing issues"; Severity = "WARNING" }
                } else {
                    $subnets[$networkId] = $adapter.Name
                }
            }
        }
    }

    # DNS reachability test (ping the primary DNS)
    $dnsReachable = "Not tested"
    foreach ($adapter in $netAdapters) {
        if ($adapter.DNS -ne "None" -and $adapter.DNS -ne "") {
            $primaryDNS = ($adapter.DNS -split ",")[0].Trim()
            try {
                $pingResult = Test-Connection -ComputerName $primaryDNS -Count 1 -Quiet -ErrorAction SilentlyContinue
                $dnsReachable = if ($pingResult) { "Yes ($primaryDNS)" } else { "No ($primaryDNS)" }
                if (-not $pingResult) {
                    $netIssues += @{ Adapter = $adapter.Name; Issue = "Primary DNS server $primaryDNS is not responding to ping"; Severity = "WARNING" }
                }
            } catch {
                $dnsReachable = "Test failed"
            }
            break
        }
    }

    # Internet connectivity test (DNS resolution)
    $internetAccess = "Not tested"
    try {
        $dnsTest = [System.Net.Dns]::GetHostAddresses("www.google.com")
        if ($dnsTest) { $internetAccess = "Yes (DNS resolving)" } else { $internetAccess = "No" }
    } catch {
        $internetAccess = "No (DNS resolution failed)"
        $netIssues += @{ Adapter = "General"; Issue = "Internet DNS resolution failed - no internet access"; Severity = "ERROR" }
    }

    # Public IP (quick external lookup)
    $publicIP = "Not available"
    try {
        $webClient = New-Object System.Net.WebClient
        $publicIP = ($webClient.DownloadString("https://api.ipify.org")).Trim()
    } catch {
        $publicIP = "Could not determine"
    }

    # Physical printers (exclude virtual/software printers)
    $physicalPrinters = @()
    try {
        $virtualPrinterPatterns = @(
            "Microsoft Print to PDF",
            "Microsoft XPS Document Writer",
            "OneNote",
            "Adobe PDF",
            "Bullzip PDF",
            "CutePDF",
            "doPDF",
            "Foxit.*PDF",
            "PDF24",
            "PDFCreator",
            "Nitro PDF",
            "Send to OneNote",
            "Fax",
            "Microsoft Shared Fax",
            "Snagit",
            "XPS",
            "Print to Evernote",
            "novaPDF",
            "PrimoPDF",
            "FreePDF",
            "PDF-XChange",
            "Wondershare PDFelement",
            "TOSHIBA e-STUDIO PDF",
            "Canon.*FAX",
            "Nul$"
        )
        $virtualPattern = ($virtualPrinterPatterns | ForEach-Object { "($_)" }) -join "|"

        $allPrinters = Get-CimInstance Win32_Printer
        foreach ($printer in $allPrinters) {
            # Skip virtual/software printers
            if ($printer.Name -match $virtualPattern) { continue }

            # Skip printers with no port or only FILE: / PORTPROMPT: ports (virtual indicators)
            $portName = $printer.PortName
            if ($portName -match "^(FILE:|PORTPROMPT:|NUL:|nul:)") { continue }

            # Accept printers with physical/network ports or flagged as network printers
            # Known physical ports: USB, LPT, WSD, TCP/IP, DOT4, IP addresses, Ne ports
            # Also accept anything the OS flags as a network printer, or HP/vendor-specific ports
            $isPhysical = $portName -match "(USB|LPT|WSD|IP_|TCPMON|DOT4|Ne\d)" -or
                          $printer.Network -or
                          $portName -match '^\d+\.\d+\.\d+\.\d+' -or
                          $portName -match '(HP|Canon|Epson|Brother|Samsung|Xerox|Ricoh|Lexmark|Kyocera)'

            if (-not $isPhysical) { continue }

            $statusText = switch ($printer.PrinterStatus) {
                1 { "Other" }
                2 { "Unknown" }
                3 { "Idle" }
                4 { "Printing" }
                5 { "Warming Up" }
                6 { "Stopped" }
                7 { "Offline" }
                default { "Unknown ($($printer.PrinterStatus))" }
            }

            # Determine connection type
            $connType = if ($printer.Network) { "Network" }
                        elseif ($portName -match "USB") { "USB" }
                        elseif ($portName -match "LPT") { "Parallel" }
                        elseif ($portName -match "WSD") { "WSD (Wi-Fi/Network)" }
                        elseif ($portName -match '^\d+\.\d+\.\d+\.\d+|TCPMON|IP_|Ne\d') { "Network (TCP/IP)" }
                        elseif ($portName -match 'HP|Canon|Epson|Brother') { "Network (Vendor)" }
                        else { "Local" }

            $isDefault = $printer.Default

            $physicalPrinters += @{
                Name       = $printer.Name
                Port       = $portName
                Status     = $statusText
                Connection = $connType
                Default    = $isDefault
                Shared     = $printer.Shared
                Driver     = $printer.DriverName
            }

            # Flag offline printers
            if ($statusText -eq "Offline" -or $statusText -eq "Stopped") {
                $netIssues += @{ Adapter = "Printer"; Issue = "Printer '$($printer.Name)' is $statusText"; Severity = "WARNING" }
            }
        }
        Write-Log "Found $($physicalPrinters.Count) physical printer(s)"
    } catch {
        Write-Log "Printer detection failed: $_" "WARNING"
    }

    # Internet speed test (using Ookla Speedtest CLI)
    $speedTest = @{
        DownloadMbps = "Not tested"
        UploadMbps   = "Not tested"
        PingMs       = "Not tested"
        JitterMs     = "Not tested"
        ServerName   = "N/A"
        ServerLocation = "N/A"
        ISP          = "N/A"
        Tested       = $false
        Error        = ""
    }

    try {
        Write-Log "Running internet speed test..."
        $toolsDir = "C:\Teamviewer\Tools"
        $speedtestDir = Join-Path $toolsDir "speedtest"
        $speedtestExe = Join-Path $speedtestDir "speedtest.exe"

        if (-not (Test-Path $toolsDir)) {
            New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null
        }

        # Download Speedtest CLI if not present
        if (-not (Test-Path $speedtestExe)) {
            Write-Log "Downloading Speedtest CLI..."
            $speedtestZip = Join-Path $toolsDir "speedtest.zip"
            $speedtestUrl = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip"
            (New-Object System.Net.WebClient).DownloadFile($speedtestUrl, $speedtestZip)

            if (-not (Test-Path $speedtestDir)) {
                New-Item -ItemType Directory -Path $speedtestDir -Force | Out-Null
            }
            Expand-Archive -Path $speedtestZip -DestinationPath $speedtestDir -Force
            Remove-Item $speedtestZip -Force -ErrorAction SilentlyContinue
            Write-Log "Speedtest CLI downloaded to $speedtestDir"
        }

        if (Test-Path $speedtestExe) {
            # Accept license and run in JSON format
            $speedtestOutput = & $speedtestExe --accept-license --accept-gdpr --format=json 2>&1
            $speedtestJson = $speedtestOutput | Out-String | ConvertFrom-Json

            if ($speedtestJson.download -and $speedtestJson.upload) {
                # Speedtest CLI reports bandwidth in bytes per second
                $dlMbps = [math]::Round($speedtestJson.download.bandwidth * 8 / 1000000, 1)
                $ulMbps = [math]::Round($speedtestJson.upload.bandwidth * 8 / 1000000, 1)
                $pingMs = [math]::Round($speedtestJson.ping.latency, 1)
                $jitterMs = [math]::Round($speedtestJson.ping.jitter, 1)

                $speedTest.DownloadMbps = $dlMbps
                $speedTest.UploadMbps = $ulMbps
                $speedTest.PingMs = $pingMs
                $speedTest.JitterMs = $jitterMs
                $speedTest.ServerName = $speedtestJson.server.name
                $speedTest.ServerLocation = "$($speedtestJson.server.location), $($speedtestJson.server.country)"
                $speedTest.ISP = $speedtestJson.isp
                $speedTest.Tested = $true

                Write-Log "Speed test complete: Down $dlMbps Mbps, Up $ulMbps Mbps, Ping $pingMs ms, Jitter $jitterMs ms"

                # Flag poor speed results
                if ($dlMbps -lt 10) {
                    $netIssues += @{ Adapter = "Internet"; Issue = "Download speed is very slow ($dlMbps Mbps) - check connection or plan"; Severity = "WARNING" }
                }
                if ($ulMbps -lt 5) {
                    $netIssues += @{ Adapter = "Internet"; Issue = "Upload speed is very slow ($ulMbps Mbps) - may affect backups and video calls"; Severity = "WARNING" }
                }
                if ($pingMs -gt 100) {
                    $netIssues += @{ Adapter = "Internet"; Issue = "High latency ($pingMs ms) - may cause lag in video calls and remote access"; Severity = "WARNING" }
                }
                if ($jitterMs -gt 30) {
                    $netIssues += @{ Adapter = "Internet"; Issue = "High jitter ($jitterMs ms) - connection is unstable, may cause audio/video stuttering"; Severity = "WARNING" }
                }
            } else {
                $speedTest.Error = "Speedtest returned no data"
                Write-Log "Speed test returned no data" "WARNING"
            }
        } else {
            $speedTest.Error = "Speedtest CLI not found after download"
            Write-Log "Speedtest CLI not found at $speedtestExe" "WARNING"
        }
    } catch {
        $speedTest.Error = $_.ToString()
        Write-Log "Speed test failed: $_" "WARNING"
    }

    $netStatus = if ($netIssues | Where-Object { $_.Severity -eq "ERROR" }) { "FAIL" }
                 elseif ($netIssues.Count -gt 0) { "WARNING" }
                 else { "PASS" }

    $Results.NetworkConfig = @{
        Status              = $netStatus
        AdapterCount        = $netAdapters.Count
        Adapters            = $netAdapters
        Issues              = $netIssues
        IssueCount          = $netIssues.Count
        DNSReachable        = $dnsReachable
        InternetAccess      = $internetAccess
        PublicIP            = $publicIP
        SpeedTest           = $speedTest
        Printers            = $physicalPrinters
        PrinterCount        = $physicalPrinters.Count
    }

    Write-Log "Network config check complete: $($netAdapters.Count) adapter(s), $($netIssues.Count) issue(s)"

} catch {
    Write-Log "Network configuration check failed: $_" "ERROR"
    $Results.NetworkConfig = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Network Config: $_"
}


# ============================================================================
# MODULE 17: ADWCLEANER
# ============================================================================
if (-not $SkipAdwCleaner) {
    Write-Log "Downloading and running AdwCleaner..."

    try {
        $adwDir = "C:\Teamviewer\Tools"
        $adwPath = Join-Path $adwDir "adwcleaner.exe"
        $adwLogDir = "C:\AdwCleaner\Logs"

        if (-not (Test-Path $adwDir)) {
            New-Item -ItemType Directory -Path $adwDir -Force | Out-Null
        }

        # Download latest AdwCleaner
        Write-Log "Downloading AdwCleaner..."
        $downloadUrl = "https://downloads.malwarebytes.com/file/adwcleaner"

        # Use BITS for reliable download, fallback to WebClient
        try {
            Start-BitsTransfer -Source $downloadUrl -Destination $adwPath -ErrorAction Stop
        } catch {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            (New-Object System.Net.WebClient).DownloadFile($downloadUrl, $adwPath)
        }

        if (Test-Path $adwPath) {
            Write-Log "AdwCleaner downloaded. Running scan (scan only, no cleaning)..."

            # Run AdwCleaner in scan-only mode
            # /eula = accept EULA, /scan = scan only (no removal), /noreboot
            $adwProcess = Start-Process -FilePath $adwPath -ArgumentList "/eula", "/scan", "/noreboot" -Wait -PassThru -NoNewWindow -ErrorAction Stop

            # Wait for log file to be written
            Start-Sleep -Seconds 5

            # Find the most recent AdwCleaner log
            $adwLogContent = "No log found"
            $detectionCount = 0

            if (Test-Path $adwLogDir) {
                $adwLog = Get-ChildItem -Path $adwLogDir -Filter "*.txt" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1

                if ($adwLog) {
                    $adwLogContent = Get-Content -LiteralPath $adwLog.FullName | Out-String

                    # Separate adware/malware from preinstalled software
                    $logLines = $adwLogContent -split "`n"
                    $adwareLines = @()
                    $preinstalledLines = @()
                    $inPreinstalled = $false
                    foreach ($line in $logLines) {
                        if ($line -match "\*+\s*\[\s*Preinstalled Software\s*\]\s*\*+") {
                            $inPreinstalled = $true
                            continue
                        }
                        if ($inPreinstalled) {
                            # Capture preinstalled entries (non-blank, non-header lines with actual content)
                            # Skip AdwCleaner log file listings (e.g. "AdwCleaner[S00].txt - [1250 octets]")
                            if ($line.Trim() -and $line -notmatch "^\*+\s*\[" -and $line -notmatch "^#" -and $line -notmatch "AdwCleaner\[.+\]\.txt") {
                                $preinstalledLines += $line.Trim()
                            }
                        } else {
                            # Actual adware/malware detections
                            if ($line -match "^(PUP\.|Adware\.|Toolbar\.|Hijack\.|Trojan\.|Spyware\.|Ransomware\.)") {
                                $adwareLines += $line
                            }
                        }
                    }
                    $detectionCount = $adwareLines.Count
                }
            }

            # Also check the JSON results if available
            $adwJsonPath = "C:\AdwCleaner\AdwCleaner_Results.json"
            if (Test-Path $adwJsonPath) {
                $adwJson = Get-Content -LiteralPath $adwJsonPath | Out-String | ConvertFrom-Json -ErrorAction SilentlyContinue
            }

            $Results.AdwCleaner = @{
                Status              = if ($detectionCount -eq 0) { "CLEAN" } else { "WARNING  - $detectionCount detection(s)" }
                ExitCode            = $adwProcess.ExitCode
                DetectionCount      = $detectionCount
                Detections          = $adwareLines
                PreinstalledCount   = $preinstalledLines.Count
                PreinstalledItems   = $preinstalledLines
                LogFile             = if ($adwLog) { $adwLog.FullName } else { "N/A" }
                LogSummary          = if ($adwLogContent.Length -gt 2000) { $adwLogContent.Substring(0, 2000) + "... [truncated]" } else { $adwLogContent }
            }

            Write-Log "AdwCleaner scan complete: $detectionCount detection(s)"
        } else {
            throw "Download failed  - file not found at $adwPath"
        }

    } catch {
        Write-Log "AdwCleaner failed: $_" "ERROR"
        $Results.AdwCleaner = @{ Status = "ERROR"; Error = $_.ToString() }
        $Results.Errors += "AdwCleaner: $_"
    }
} else {
    Write-Log "AdwCleaner skipped (SkipAdwCleaner flag set)"
    $Results.AdwCleaner = @{ Status = "SKIPPED" }
}


# ============================================================================
# MODULE 18: SYSTEM RESTORE POINT AUDIT
# ============================================================================
Write-Log "Checking system restore points..."

try {
    $restorePoints = @()
    try {
        $restorePoints = @(Get-ComputerRestorePoint -ErrorAction Stop)
    } catch {
        # Get-ComputerRestorePoint may fail if System Restore is disabled
    }

    # Check if System Restore is enabled on the OS drive
    $osDrive = $env:SystemDrive  # typically C:
    $srEnabled = $false
    try {
        $sr = Get-CimInstance -ClassName SystemRestoreConfig -Namespace "root\DEFAULT" -ErrorAction Stop |
              Where-Object { $_.DiskNumber -eq 0 -or $_.RPSessionInterval -ge 0 }
        # Also check via vssadmin or registry
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
        $rpDisabled = (Get-ItemProperty -Path $regPath -Name "RPSessionInterval" -ErrorAction SilentlyContinue).RPSessionInterval
        $srDisabledReg = (Get-ItemProperty -Path $regPath -Name "DisableSR" -ErrorAction SilentlyContinue).DisableSR
        $srEnabled = ($srDisabledReg -ne 1)
    } catch {
        # Fallback: check registry directly
        try {
            $srDisabledReg = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "DisableSR" -ErrorAction Stop).DisableSR
            $srEnabled = ($srDisabledReg -ne 1)
        } catch {
            $srEnabled = $false
        }
    }

    # Check allocated disk space for restore points
    $allocatedGB = "Unknown"
    $shadowStorageConfigured = $false
    try {
        $vssOutput = vssadmin list shadowstorage 2>&1 | Out-String
        if ($vssOutput -match "For volume: \(C:\)") {
            $shadowStorageConfigured = $true
        }
        if ($vssOutput -match "Used Shadow Copy Storage space:\s*([\d.,]+)\s*(MB|GB|TB)") {
            $usedVal = [double]($Matches[1] -replace ',', '')
            $usedUnit = $Matches[2]
            $allocatedGB = switch ($usedUnit) {
                "MB" { "$([math]::Round($usedVal / 1024, 2)) GB" }
                "GB" { "$([math]::Round($usedVal, 2)) GB" }
                "TB" { "$([math]::Round($usedVal * 1024, 2)) GB" }
                default { "$usedVal $usedUnit" }
            }
        }
    } catch {}

    $rpCount = $restorePoints.Count
    $newestRP = $null
    $oldestRP = $null
    $daysSinceLatest = -1

    if ($rpCount -gt 0) {
        $newestRP = ($restorePoints | Sort-Object CreationTime -Descending | Select-Object -First 1)
        $oldestRP = ($restorePoints | Sort-Object CreationTime | Select-Object -First 1)
        $daysSinceLatest = [math]::Floor(((Get-Date) - $newestRP.CreationTime).TotalDays)
    }

    # Determine status
    $rpStatus = "PASS"
    $rpNote = ""
    if (-not $srEnabled) {
        $rpStatus = "WARNING"
        $rpNote = "System Restore is DISABLED on this machine"
    } elseif ($rpCount -eq 0 -and $shadowStorageConfigured) {
        # Shadow storage is configured but Get-ComputerRestorePoint returns nothing.
        # On Windows 11 24H2+ builds, Checkpoint-Computer and Get-ComputerRestorePoint
        # are broken (srservice removed), but System Protection still works automatically.
        $rpStatus = "PASS"
        $rpNote = "System Protection enabled (shadow storage configured). Restore points are managed automatically by Windows."
    } elseif ($rpCount -eq 0) {
        $rpStatus = "WARNING"
        $rpNote = "No restore points found - recommend creating one"
    } elseif ($daysSinceLatest -gt 30) {
        $rpStatus = "WARNING"
        $rpNote = "Latest restore point is $daysSinceLatest days old"
    } else {
        $rpNote = "Latest restore point is $daysSinceLatest day(s) old"
    }

    $Results.RestorePoints = @{
        Status           = $rpStatus
        Enabled          = $srEnabled
        Count            = $rpCount
        DiskUsage        = $allocatedGB
        DaysSinceLatest  = $daysSinceLatest
        Note             = $rpNote
        Points           = @()
    }

    # Store up to 5 most recent restore points for the report
    if ($rpCount -gt 0) {
        $Results.RestorePoints.NewestDate = $newestRP.CreationTime.ToString("d MMM yyyy HH:mm")
        $Results.RestorePoints.OldestDate = $oldestRP.CreationTime.ToString("d MMM yyyy HH:mm")

        $recentPoints = $restorePoints | Sort-Object CreationTime -Descending | Select-Object -First 5
        foreach ($rp in $recentPoints) {
            $Results.RestorePoints.Points += @{
                Description = $rp.Description
                Created     = $rp.CreationTime.ToString("d MMM yyyy HH:mm")
                Type        = switch ($rp.RestorePointType) {
                    0  { "Application Install" }
                    1  { "Application Uninstall" }
                    6  { "Restore" }
                    7  { "Checkpoint" }
                    10 { "Device Driver Install" }
                    12 { "Modify Settings" }
                    13 { "Cancel-Restore" }
                    default { "Type $($rp.RestorePointType)" }
                }
            }
        }
    }

    Write-Log "Restore Points: $rpCount found, System Restore $(if ($srEnabled) {'enabled'} else {'DISABLED'}), status: $rpStatus"

} catch {
    Write-Log "Restore point audit failed: $_" "ERROR"
    $Results.RestorePoints = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Restore Points: $_"
}


# ============================================================================
# MODULE 19: TELEMETRY SERVICES AUDIT
# Reports on Windows telemetry/data-collection services - read-only, no changes
# ============================================================================
Write-Log "Auditing telemetry services..."

try {
    $telemetryServices = @(
        @{ Name = "DiagTrack";             Friendly = "Connected User Experiences and Telemetry" }
        @{ Name = "dmwappushservice";       Friendly = "Device Management WAP Push" }
        @{ Name = "WerSvc";                Friendly = "Windows Error Reporting" }
        @{ Name = "WMPNetworkSvc";         Friendly = "Windows Media Player Network Sharing" }
        @{ Name = "diagnosticshub.standardcollector.service"; Friendly = "Diagnostics Hub Collector" }
        @{ Name = "lfsvc";                 Friendly = "Geolocation Service" }
        @{ Name = "MapsBroker";            Friendly = "Downloaded Maps Manager" }
    )

    $telResults = @()
    $runningCount = 0
    $disabledCount = 0

    foreach ($ts in $telemetryServices) {
        $svc = Get-Service -Name $ts.Name -ErrorAction SilentlyContinue
        if ($svc) {
            $startType = "Unknown"
            try {
                $startType = (Get-CimInstance Win32_Service -Filter "Name='$($ts.Name)'" -ErrorAction Stop).StartMode
            } catch {}

            $telResults += @{
                ServiceName = $ts.Name
                FriendlyName = $ts.Friendly
                Status       = $svc.Status.ToString()
                StartType    = $startType
            }

            if ($svc.Status -eq "Running") { $runningCount++ }
            if ($startType -eq "Disabled") { $disabledCount++ }
        }
        # If service not found, skip it (not all are present on every Windows edition)
    }

    $totalFound = $telResults.Count
    $telStatus = if ($runningCount -eq 0 -and $totalFound -gt 0) { "PASS" }
                 elseif ($runningCount -le 2) { "PASS" }
                 else { "INFO" }

    $Results.TelemetryServices = @{
        Status        = $telStatus
        TotalFound    = $totalFound
        RunningCount  = $runningCount
        DisabledCount = $disabledCount
        Services      = $telResults
    }

    Write-Log "Telemetry audit: $totalFound services found, $runningCount running, $disabledCount disabled"

} catch {
    Write-Log "Telemetry services audit failed: $_" "ERROR"
    $Results.TelemetryServices = @{ Status = "ERROR"; Error = $_.ToString() }
    $Results.Errors += "Telemetry Services: $_"
}


# ============================================================================
# REPORT GENERATION
# ============================================================================
Write-Log "Generating HTML report..."

$overallStatus = "PASS"
$warningCount = 0
$errorCount = 0

$checkModules = @($Results.SFC, $Results.DiskHealth, $Results.WindowsUpdates, $Results.iDriveBackup, $Results.Malwarebytes, $Results.Defender, $Results.EventLogErrors, $Results.PendingReboot, $Results.Firewall, $Results.UserAccounts, $Results.DISM, $Results.TempFiles, $Results.StartupPrograms, $Results.ScheduledTasks, $Results.BrowserExtensions, $Results.ServiceStatus, $Results.NetworkConfig, $Results.AdwCleaner, $Results.RestorePoints, $Results.TelemetryServices)
foreach ($module in $checkModules) {
    if ($module.Status -match "ERROR|FAIL") { $errorCount++; $overallStatus = "FAIL" }
    elseif ($module.Status -match "WARNING") { $warningCount++; if ($overallStatus -ne "FAIL") { $overallStatus = "WARNING" } }
}

# Status colour helper
function Get-StatusBadge {
    param([string]$Status)
    $colour = switch -Regex ($Status) {
        "PASS|CLEAN|UP TO DATE|REPAIRED" { "#22c55e" }  # green
        "WARNING|UPDATES AVAILABLE|CHECK REQUIRED|NOT INSTALLED" { "#f59e0b" }  # amber
        "FAIL|ERROR" { "#ef4444" }  # red
        "SKIPPED" { "#94a3b8" }  # grey
        default { "#64748b" }
    }
    return "<span style='background:$colour;color:white;padding:3px 10px;border-radius:4px;font-weight:bold;font-size:0.85em;'>$Status</span>"
}

$overallBadge = Get-StatusBadge $overallStatus

# Build the HTML
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Maintenance Report  - $ComputerName</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f1f5f9; color: #1e293b; line-height: 1.5; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        .header { background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%); color: white; padding: 30px; border-radius: 12px 12px 0 0; }
        .header h1 { font-size: 1.5em; margin-bottom: 5px; }
        .header .subtitle { opacity: 0.8; font-size: 0.9em; }
        .header .overall { margin-top: 15px; font-size: 1.1em; }
        .section { background: white; border: 1px solid #e2e8f0; margin-bottom: 2px; padding: 20px 25px; }
        .section:last-child { border-radius: 0 0 12px 12px; margin-bottom: 20px; }
        .section h2 { font-size: 1.1em; color: #334155; margin-bottom: 10px; display: flex; align-items: center; gap: 10px; }
        .section h2 .icon { font-size: 1.3em; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 0.9em; }
        th { background: #f8fafc; text-align: left; padding: 8px 12px; border-bottom: 2px solid #e2e8f0; color: #64748b; font-weight: 600; }
        td { padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }
        tr:hover td { background: #f8fafc; }
        .flagged { background: #fef2f2 !important; }
        .detail { color: #64748b; font-size: 0.85em; }
        .warning-text { color: #d97706; font-weight: 600; }
        .error-text { color: #dc2626; font-weight: 600; }
        .footer { text-align: center; color: #94a3b8; font-size: 0.8em; padding: 15px; }
        pre { background: #f8fafc; padding: 12px; border-radius: 6px; font-size: 0.8em; overflow-x: auto; max-height: 200px; overflow-y: auto; }
    </style>
</head>
<body>
<div class="container">

<div class="header">
    <h1>PC Masterclass  - Maintenance Report</h1>
    <div class="subtitle">$(if ($ClientName) { "$ClientName &bull; " })$ComputerName &bull; $($Results.OSVersion) &bull; $(Get-Date -Format 'dddd, d MMMM yyyy h:mm tt')</div>
    <div class="overall">Overall Status: $overallBadge &nbsp; ($warningCount warning(s), $errorCount error(s))</div>
</div>

<!-- SYSTEM INFO -->
<div class="section">
    <h2><span class="icon">&#x1F4BB;</span> System Information</h2>
    <table>
        $(if ($ClientName) { "<tr><td><strong>Client</strong></td><td><strong style='font-size:1.1em;'>$ClientName</strong></td></tr>" })
        <tr><td><strong>Computer</strong></td><td>$ComputerName</td></tr>
        <tr><td><strong>Make / Model</strong></td><td>$($Results.SystemInfo.Manufacturer) $($Results.SystemInfo.Model)</td></tr>
        <tr><td><strong>Serial Number</strong></td><td>$($Results.SystemInfo.SerialNumber)</td></tr>
        <tr><td><strong>OS</strong></td><td>$($Results.SystemInfo.OS) - Build $($Results.SystemInfo.Build)</td></tr>
        <tr><td><strong>CPU</strong></td><td>$($Results.SystemInfo.CPU) ($($Results.SystemInfo.CPUCores) cores / $($Results.SystemInfo.CPUThreads) threads @ $($Results.SystemInfo.CPUMaxSpeedGHz) GHz)$(
            if ($Results.SystemInfo.CPUBenchmark.Score -gt 0) {
                " &mdash; <span style='color:$($Results.SystemInfo.CPUBenchmark.TierColor);font-weight:bold;'>$($Results.SystemInfo.CPUBenchmark.Tier)</span> <span style='color:#6c757d;font-size:0.9em;'>(PassMark: $($Results.SystemInfo.CPUBenchmark.Score.ToString('N0')) &bull; defs: $($Results.SystemInfo.CPUBenchmark.DefinitionsAge))</span>"
            }
        )</td></tr>
        <tr><td><strong>RAM</strong></td><td>$($Results.SystemInfo.TotalRAM_GB) GB</td></tr>
        <tr><td><strong>BIOS</strong></td><td>$($Results.SystemInfo.BIOSVersion)</td></tr>
        <tr><td><strong>Last Boot</strong></td><td>$($Results.SystemInfo.LastBoot) (Uptime: $($Results.SystemInfo.UptimeDays) days)</td></tr>
        <tr><td><strong>Script Version</strong></td><td>v$ScriptVersion</td></tr>
    </table>
"@

# BitLocker section
$html += "<h3 style='margin-top:12px;margin-bottom:6px;'>BitLocker Encryption</h3><table>"
foreach ($blVol in $Results.SystemInfo.BitLocker) {
    $blColor = if ($blVol.ProtectionStatus -eq "On") { "#28a745" } elseif ($blVol.ProtectionStatus -eq "Off") { "#dc3545" } else { "#6c757d" }
    $html += "<tr><td><strong>$($blVol.Drive)</strong></td><td><span style='color:$blColor;font-weight:bold;'>$($blVol.ProtectionStatus)</span>"
    if ($blVol.EncryptionMethod -ne "N/A") {
        $html += " ($($blVol.EncryptionMethod) - $($blVol.VolumeStatus))"
    } else {
        $html += " - $($blVol.ProtectionStatus)"
    }
    $html += "</td></tr>"
}
$html += "</table>"

# Network adapters section
$html += "<h3 style='margin-top:12px;margin-bottom:6px;'>Network Adapters</h3><table>"
$html += "<tr style='background:#f8f9fa;'><td><strong>Adapter</strong></td><td><strong>IPv4</strong></td><td><strong>Gateway</strong></td><td><strong>DNS</strong></td><td><strong>MAC</strong></td><td><strong>DHCP</strong></td></tr>"
foreach ($nic in $Results.SystemInfo.NetworkAdapters) {
    $dhcpText = if ($nic.DHCP) { "Yes" } else { "Static" }
    $html += "<tr><td>$($nic.Name)</td><td>$($nic.IPv4)</td><td>$($nic.Gateway)</td><td>$($nic.DNS)</td><td>$($nic.MAC)</td><td>$dhcpText</td></tr>"
}
$html += "</table>"

# Temperatures section
$html += "<h3 style='margin-top:12px;margin-bottom:6px;'>Temperatures</h3>"
if ($Results.SystemInfo.Temperatures[0].Zone -ne "N/A") {
    $html += "<table>"
    $html += "<tr style='background:#f8f9fa;'><td><strong>Zone</strong></td><td><strong>Temp (C)</strong></td><td><strong>Status</strong></td></tr>"
    foreach ($temp in $Results.SystemInfo.Temperatures) {
        $tempColor = if ($temp.TempC -gt 85) { "#dc3545" } elseif ($temp.TempC -gt 70) { "#fd7e14" } else { "#28a745" }
        $tempLabel = if ($temp.TempC -gt 85) { "HOT" } elseif ($temp.TempC -gt 70) { "Warm" } else { "Normal" }
        $html += "<tr><td>$($temp.Zone)</td><td>$($temp.TempC)</td><td><span style='color:$tempColor;font-weight:bold;'>$tempLabel</span></td></tr>"
    }
    $html += "</table>"
} else {
    $html += "<p class='detail'>Temperature sensors not accessible via WMI on this system.</p>"
}

$html += @"
</div>

<!-- SFC -->
<div class="section">
    <h2><span class="icon">&#x1F6E1;</span> System File Integrity (SFC) $(Get-StatusBadge $Results.SFC.Status)</h2>
"@

if ($Results.SFC.Status -eq "SKIPPED") {
    $html += "<p class='detail'>Skipped by configuration.</p>"
} elseif ($Results.SFC.Status -eq "ERROR") {
    $html += "<p class='error-text'>Error: $($Results.SFC.Error)</p>"
} else {
    $html += @"
    <p>Scan completed in $($Results.SFC.DurationMinutes) minutes. CBS log entries with corruption references: $($Results.SFC.CorruptEntries)</p>
"@
}

$html += @"
</div>

<!-- DISK HEALTH -->
<div class="section">
    <h2><span class="icon">&#x1F4BE;</span> Disk Health $(Get-StatusBadge $Results.DiskHealth.Status)</h2>
"@

if ($Results.DiskHealth.Disks) {
    $html += "<table><tr><th>Disk</th><th>Type</th><th>Size</th><th>Health</th><th>Temp</th></tr>"
    foreach ($disk in $Results.DiskHealth.Disks) {
        $html += "<tr><td>$($disk.Name)</td><td>$($disk.Type)</td><td>$($disk.SizeGB) GB</td><td>$($disk.HealthStatus)</td><td>$($disk.Temperature)</td></tr>"
    }
    $html += "</table>"
}

if ($Results.DiskHealth.Volumes) {
    $html += "<table><tr><th>Volume</th><th>Label</th><th>Total</th><th>Free</th><th>Free %</th></tr>"
    foreach ($vol in $Results.DiskHealth.Volumes) {
        $rowClass = if ($vol.Warning) { "class='flagged'" } else { "" }
        $freeText = if ($vol.Warning) { "<span class='warning-text'>$($vol.FreePercent)% WARNING:</span>" } else { "$($vol.FreePercent)%" }
        $html += "<tr $rowClass><td>$($vol.Drive)</td><td>$($vol.Label)</td><td>$($vol.SizeGB) GB</td><td>$($vol.FreeGB) GB</td><td>$freeText</td></tr>"
    }
    $html += "</table>"
}

$html += @"
</div>

<!-- WINDOWS UPDATES -->
<div class="section">
    <h2><span class="icon">&#x1F504;</span> Windows Updates $(Get-StatusBadge $Results.WindowsUpdates.Status)</h2>
"@

if ($Results.WindowsUpdates.Status -eq "SKIPPED") {
    $html += "<p class='detail'>Skipped by configuration.</p>"
} elseif ($Results.WindowsUpdates.PendingCount -eq 0) {
    $html += "<p>No pending updates. System is up to date.</p>"
} else {
    $html += "<p>$($Results.WindowsUpdates.PendingCount) update(s) pending. Install result: $($Results.WindowsUpdates.InstallResult)</p>"
    $html += "<table><tr><th>Update</th><th>KB</th><th>Size</th><th>Critical?</th></tr>"
    foreach ($upd in $Results.WindowsUpdates.Updates) {
        $critBadge = if ($upd.IsImportant) { "<span class='warning-text'>Yes</span>" } else { "No" }
        $html += "<tr><td>$($upd.Title)</td><td>$($upd.KB)</td><td>$($upd.SizeMB) MB</td><td>$critBadge</td></tr>"
    }
    $html += "</table>"
}

$html += @"
</div>

<!-- iDRIVE BACKUP -->
<div class="section">
    <h2><span class="icon">&#x2601;</span> iDrive Backup $(Get-StatusBadge $Results.iDriveBackup.Status)</h2>
"@

if (-not $Results.iDriveBackup.Installed) {
    $html += "<p class='detail'>iDrive client not found on this system.</p>"
} else {
    $svcRunning = if ($Results.iDriveBackup.ServiceRunning) { "Running" } else { "<span class='warning-text'>Not Running</span>" }
    $html += @"
    <table>
        <tr><td><strong>Service Status</strong></td><td>$svcRunning</td></tr>
        <tr><td><strong>Last Backup</strong></td><td>$($Results.iDriveBackup.LastBackupDate)</td></tr>
        <tr><td><strong>Backup Result</strong></td><td>$($Results.iDriveBackup.LastBackupResult)</td></tr>
        <tr><td><strong>Files Backed Up</strong></td><td>$($Results.iDriveBackup.FilesBackedUp) of $($Results.iDriveBackup.FilesTotal) considered</td></tr>
        <tr><td><strong>Files In Sync</strong></td><td>$($Results.iDriveBackup.FilesInSync)</td></tr>
        <tr><td><strong>Duration</strong></td><td>$($Results.iDriveBackup.Duration)</td></tr>
        <tr><td><strong>Computer Name</strong></td><td>$($Results.iDriveBackup.ComputerName)</td></tr>
        <tr><td><strong>Account</strong></td><td>$($Results.iDriveBackup.Account)</td></tr>
    </table>
"@
}

$html += @"
</div>

<!-- MALWAREBYTES -->
<div class="section">
    <h2><span class="icon">&#x1F6E1;</span> Malwarebytes Endpoint Protection $(Get-StatusBadge $Results.Malwarebytes.Status)</h2>
"@

if (-not $Results.Malwarebytes.Installed) {
    $html += "<p class='detail'>Malwarebytes not found on this system.</p>"
} else {
    $mbSvcRunning = if ($Results.Malwarebytes.ServiceRunning) { "Running" } else { "<span class='warning-text'>Not Running</span>" }
    $mbRtColor = if ($Results.Malwarebytes.RealTimeProtection -eq "Enabled") { "color:#28a745" } elseif ($Results.Malwarebytes.RealTimeProtection -eq "Disabled") { "color:#dc3545" } else { "" }
    $html += @"
    <table>
        <tr><td><strong>Product</strong></td><td>$($Results.Malwarebytes.ProductType)</td></tr>
        <tr><td><strong>Service Status</strong></td><td>$mbSvcRunning</td></tr>
        <tr><td><strong>Real-Time Protection</strong></td><td><span style='$mbRtColor;font-weight:bold;'>$($Results.Malwarebytes.RealTimeProtection)</span></td></tr>
    </table>
"@
}

$html += @"
</div>

<!-- WINDOWS DEFENDER -->
<div class="section">
    <h2><span class="icon">&#x1F6E1;</span> Windows Defender / Antivirus $(Get-StatusBadge $Results.Defender.Status)</h2>
"@

if ($Results.Defender.ThirdPartyAV) {
    $html += @"
    <table>
        <tr><td><strong>Third-Party Antivirus</strong></td><td>$($Results.Defender.ThirdPartyAV)</td></tr>
        <tr><td><strong>Real-Time Protection</strong></td><td>$($Results.Defender.RealTimeProtection)</td></tr>
    </table>
"@
} elseif ($Results.Defender.RealTimeProtection) {
    $rtClass = if ($Results.Defender.RealTimeProtection -eq "DISABLED") { "class='error-text'" } else { "" }
    $html += @"
    <table>
        <tr><td><strong>Real-Time Protection</strong></td><td $rtClass>$($Results.Defender.RealTimeProtection)</td></tr>
        <tr><td><strong>Definitions Updated</strong></td><td>$($Results.Defender.DefinitionsUpdated) ($($Results.Defender.DefinitionsAge))</td></tr>
        <tr><td><strong>Last Scan</strong></td><td>$($Results.Defender.LastScan) ($($Results.Defender.LastScanType))</td></tr>
        <tr><td><strong>Threats (Last 30 Days)</strong></td><td>$($Results.Defender.ThreatsLast30Days)</td></tr>
        <tr><td><strong>Engine Version</strong></td><td>$($Results.Defender.EngineVersion)</td></tr>
    </table>
"@
    if ($Results.Defender.ThreatsLast30Days -gt 0 -and $Results.Defender.RecentThreats) {
        $html += "<h3 style='color:#d97706;margin:10px 0 5px;font-size:0.95em;'>Recent Threats</h3>"
        $html += "<table><tr><th>Threat</th><th>Detected</th><th>Resolved</th></tr>"
        foreach ($threat in $Results.Defender.RecentThreats) {
            $resolvedText = if ($threat.Action) { "Yes" } else { "<span class='warning-text'>No</span>" }
            $html += "<tr><td>$($threat.Name)</td><td>$($threat.Date)</td><td>$resolvedText</td></tr>"
        }
        $html += "</table>"
    }
} else {
    $html += "<p class='detail'>Could not retrieve antivirus status.</p>"
}

$html += @"
</div>

<!-- EVENT LOG ERRORS -->
<div class="section">
    <h2><span class="icon">&#x1F4CB;</span> Event Log Errors (Last $($Results.EventLogErrors.HoursChecked)h) $(Get-StatusBadge $Results.EventLogErrors.Status)</h2>
"@

if ($Results.EventLogErrors.TotalEvents -eq 0) {
    $html += "<p>No critical or error events in the last $($Results.EventLogErrors.HoursChecked) hours. All clear.</p>"
} else {
    $html += "<p>$($Results.EventLogErrors.TotalEvents) event(s) found: <strong>$($Results.EventLogErrors.ActionableCount) need attention</strong>, $($Results.EventLogErrors.RoutineCount) routine (safe to ignore).</p>"

    # Show actionable events first (these need attention)
    if ($Results.EventLogErrors.ActionableEvents -and $Results.EventLogErrors.ActionableCount -gt 0) {
        $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Needs Attention ($($Results.EventLogErrors.ActionableCount))</h3>"
        $html += "<table><tr><th>Time</th><th>Level</th><th>Source</th><th>ID</th><th>What This Means</th></tr>"
        foreach ($evt in $Results.EventLogErrors.ActionableEvents) {
            $levelClass = if ($evt.Level -eq "Critical") { "class='error-text'" } else { "" }
            $html += "<tr class='flagged'><td style='white-space:nowrap'>$($evt.Time)</td><td $levelClass><strong>$($evt.Level)</strong></td><td>$($evt.Source)</td><td>$($evt.EventID)</td><td>$($evt.Note)</td></tr>"
        }
        $html += "</table>"
    }

    # Show routine events in a collapsible section
    if ($Results.EventLogErrors.RoutineEvents -and $Results.EventLogErrors.RoutineCount -gt 0) {
        $html += "<details><summary style='cursor:pointer;color:#3b82f6;margin-top:10px;'>Routine events - safe to ignore ($($Results.EventLogErrors.RoutineCount))</summary>"
        $html += "<table><tr><th>Time</th><th>Source</th><th>ID</th><th>Explanation</th></tr>"
        foreach ($evt in $Results.EventLogErrors.RoutineEvents) {
            $html += "<tr><td style='white-space:nowrap'>$($evt.Time)</td><td>$($evt.Source)</td><td>$($evt.EventID)</td><td class='detail'>$($evt.Note)</td></tr>"
        }
        $html += "</table></details>"
    }
}

$html += @"
</div>

<!-- PENDING REBOOT -->
<div class="section">
    <h2><span class="icon">&#x1F504;</span> Pending Reboot $(Get-StatusBadge $Results.PendingReboot.Status)</h2>
"@

if ($Results.PendingReboot.RebootRequired) {
    $html += "<p class='warning-text'>Reboot is pending.</p>"
    $html += "<table><tr><td><strong>Reasons</strong></td><td>$($Results.PendingReboot.Reasons)</td></tr>"
    $html += "<tr><td><strong>Uptime</strong></td><td>$($Results.PendingReboot.UptimeDays) days</td></tr></table>"
} elseif ($Results.PendingReboot.UptimeWarning) {
    $html += "<p class='warning-text'>No reboot in $($Results.PendingReboot.UptimeDays) days. Consider scheduling a restart.</p>"
} else {
    $html += "<p>No reboot pending. Uptime: $($Results.PendingReboot.UptimeDays) days.</p>"
}

$html += @"
</div>

<!-- WINDOWS FIREWALL -->
<div class="section">
    <h2><span class="icon">&#x1F6E1;</span> Windows Firewall $(Get-StatusBadge $Results.Firewall.Status)</h2>
"@

if ($Results.Firewall.Profiles) {
    $html += "<table><tr><th>Profile</th><th>Enabled</th><th>Inbound Default</th><th>Outbound Default</th></tr>"
    foreach ($fwp in $Results.Firewall.Profiles) {
        $enabledText = if ($fwp.Enabled) { "Yes" } else { "<span class='error-text'>NO</span>" }
        $html += "<tr><td>$($fwp.Profile)</td><td>$enabledText</td><td>$($fwp.DefaultInboundAction)</td><td>$($fwp.DefaultOutboundAction)</td></tr>"
    }
    $html += "</table>"
} else {
    $html += "<p class='detail'>Could not retrieve firewall status.</p>"
}

$html += @"
</div>

<!-- USER ACCOUNTS -->
<div class="section">
    <h2><span class="icon">&#x1F464;</span> User Accounts $(Get-StatusBadge $Results.UserAccounts.Status)</h2>
    <p>$($Results.UserAccounts.EnabledCount) enabled account(s), $($Results.UserAccounts.AdminCount) with admin rights.</p>
"@

if ($Results.UserAccounts.FlaggedCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Flagged Accounts ($($Results.UserAccounts.FlaggedCount))</h3>"
    $html += "<table><tr><th>Account</th><th>Admin</th><th>Last Logon</th><th>Concern</th></tr>"
    foreach ($acct in $Results.UserAccounts.FlaggedAccounts) {
        $html += "<tr class='flagged'><td><strong>$($acct.Name)</strong></td><td>$($acct.IsAdmin)</td><td>$($acct.LastLogon)</td><td class='warning-text'>$($acct.FlagReason)</td></tr>"
    }
    $html += "</table>"
}

if ($Results.UserAccounts.AdminAccounts) {
    $html += "<h3 style='margin:10px 0 5px;font-size:0.95em;'>Administrator Accounts</h3>"
    $html += "<table><tr><th>Account</th><th>Enabled</th><th>Last Logon</th><th>Password Set</th></tr>"
    foreach ($acct in $Results.UserAccounts.AdminAccounts) {
        $html += "<tr><td>$($acct.Name)</td><td>$($acct.Enabled)</td><td>$($acct.LastLogon)</td><td>$($acct.PasswordSet)</td></tr>"
    }
    $html += "</table>"
}

$html += @"
</div>

<!-- DISM HEALTH -->
<div class="section">
    <h2><span class="icon">&#x1F527;</span> DISM Component Store $(Get-StatusBadge $Results.DISM.Status)</h2>
    <p>$($Results.DISM.Result)</p>
</div>

<!-- TEMP FILES -->
<div class="section">
    <h2><span class="icon">&#x1F5D1;</span> Temp and Cache Files $(Get-StatusBadge $Results.TempFiles.Status)</h2>
    <p>Total: $($Results.TempFiles.TotalSizeGB) GB across $($Results.TempFiles.Locations.Count) location(s).</p>
"@

if ($Results.TempFiles.Locations -and $Results.TempFiles.Locations.Count -gt 0) {
    $html += "<table><tr><th>Location</th><th>Size (MB)</th><th>Files</th></tr>"
    foreach ($loc in $Results.TempFiles.Locations) {
        $sizeClass = if ($loc.SizeMB -gt 1024) { "class='warning-text'" } else { "" }
        $fileText = if ($loc.FileCount -gt 0) { "$($loc.FileCount)" } else { "-" }
        $html += "<tr><td>$($loc.Location)</td><td $sizeClass>$($loc.SizeMB)</td><td>$fileText</td></tr>"
    }
    $html += "</table>"
}

$html += @"
</div>

<!-- STARTUP PROGRAMS -->
<div class="section">
    <h2><span class="icon">&#x1F680;</span> Startup Programs $(Get-StatusBadge $Results.StartupPrograms.Status)</h2>
    <p>$($Results.StartupPrograms.TotalCount) startup items found. $($Results.StartupPrograms.FlaggedCount) flagged for review.</p>
"@

if ($Results.StartupPrograms.FlaggedCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>WARNING: Flagged Items</h3>"
    $html += "<table><tr><th>Name</th><th>Command</th><th>Source</th><th>Reason</th></tr>"
    foreach ($item in $Results.StartupPrograms.FlaggedItems) {
        $cmdTruncated = if ($item.Command.Length -gt 80) { $item.Command.Substring(0, 80) + "..." } else { $item.Command }
        $html += "<tr class='flagged'><td><strong>$($item.Name)</strong></td><td class='detail'>$cmdTruncated</td><td>$($item.Source)</td><td class='warning-text'>$($item.Reasons)</td></tr>"
    }
    $html += "</table>"
}

$html += "<details><summary style='cursor:pointer;color:#3b82f6;margin-top:10px;'>View all startup items</summary>"
$html += "<table><tr><th>Name</th><th>Command</th><th>Source</th></tr>"
foreach ($item in $Results.StartupPrograms.AllItems) {
    $cmdTruncated = if ($item.Command.Length -gt 80) { $item.Command.Substring(0, 80) + "..." } else { $item.Command }
    $html += "<tr><td>$($item.Name)</td><td class='detail'>$cmdTruncated</td><td>$($item.Source)</td></tr>"
}
$html += "</table></details>"

$html += @"
</div>

<!-- SCHEDULED TASKS AUDIT -->
<div class="section">
    <h2><span class="icon">&#x23F0;</span> Scheduled Tasks Audit $(Get-StatusBadge $Results.ScheduledTasks.Status)</h2>
    <p>$($Results.ScheduledTasks.ThirdPartyCount) third-party task(s) found out of $($Results.ScheduledTasks.TotalTasks) total. $($Results.ScheduledTasks.FlaggedCount) flagged for review.</p>
"@

if ($Results.ScheduledTasks.FlaggedCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Flagged Tasks ($($Results.ScheduledTasks.FlaggedCount))</h3>"
    $html += "<table><tr><th>Task Name</th><th>Command</th><th>Concern</th></tr>"
    foreach ($task in $Results.ScheduledTasks.FlaggedTasks) {
        $cmdTruncated = if ($task.Command.Length -gt 80) { $task.Command.Substring(0, 80) + "..." } else { $task.Command }
        $html += "<tr class='flagged'><td><strong>$($task.TaskName)</strong></td><td class='detail'>$cmdTruncated</td><td class='error-text'>$($task.FlagReason)</td></tr>"
    }
    $html += "</table>"
}

if ($Results.ScheduledTasks.ThirdPartyTasks -and $Results.ScheduledTasks.ThirdPartyCount -gt 0) {
    $html += "<details><summary style='cursor:pointer;color:#3b82f6;margin-top:10px;'>View all third-party tasks ($($Results.ScheduledTasks.ThirdPartyCount))</summary>"
    $html += "<table><tr><th>Task Name</th><th>Publisher</th><th>Status</th><th>Next Run</th></tr>"
    foreach ($task in $Results.ScheduledTasks.ThirdPartyTasks) {
        $html += "<tr><td>$($task.TaskName)</td><td>$($task.Publisher)</td><td>$($task.State)</td><td>$($task.NextRun)</td></tr>"
    }
    $html += "</table></details>"
}

$html += @"
</div>

<!-- BROWSER EXTENSION AUDIT -->
<div class="section">
    <h2><span class="icon">&#x1F310;</span> Browser Extension Audit $(Get-StatusBadge $Results.BrowserExtensions.Status)</h2>
    <p>Browsers found: $($Results.BrowserExtensions.BrowsersFound -join ', '). $($Results.BrowserExtensions.TotalExtensions) extension(s), $($Results.BrowserExtensions.FlaggedCount) flagged for review.</p>
"@

if ($Results.BrowserExtensions.FlaggedCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Flagged Extensions ($($Results.BrowserExtensions.FlaggedCount))</h3>"
    $html += "<table><tr><th>Extension</th><th>Browser</th><th>Version</th><th>Concern</th></tr>"
    foreach ($ext in $Results.BrowserExtensions.FlaggedExtensions) {
        $html += "<tr class='flagged'><td><strong>$($ext.Name)</strong></td><td>$($ext.Browser)</td><td>$($ext.Version)</td><td class='error-text'>$($ext.FlagReason)</td></tr>"
    }
    $html += "</table>"
}

if ($Results.BrowserExtensions.AllExtensions -and $Results.BrowserExtensions.TotalExtensions -gt 0) {
    $html += "<details><summary style='cursor:pointer;color:#3b82f6;margin-top:10px;'>View all extensions ($($Results.BrowserExtensions.TotalExtensions))</summary>"
    $html += "<table><tr><th>Extension</th><th>Browser</th><th>Version</th><th>Status</th></tr>"
    foreach ($ext in $Results.BrowserExtensions.AllExtensions) {
        $statusText = if ($ext.Flagged) { "<span class='error-text'>FLAGGED</span>" } else { "OK" }
        $html += "<tr><td>$($ext.Name)</td><td>$($ext.Browser)</td><td>$($ext.Version)</td><td>$statusText</td></tr>"
    }
    $html += "</table></details>"
}

$html += @"
</div>

<!-- SERVICE STATUS MONITOR -->
<div class="section">
    <h2><span class="icon">&#x2699;</span> Service Status Monitor $(Get-StatusBadge $Results.ServiceStatus.Status)</h2>
    <p>$($Results.ServiceStatus.TotalChecked) services checked. $($Results.ServiceStatus.RunningCount) running, $($Results.ServiceStatus.StoppedUnexpectedCount) stopped unexpectedly.</p>
"@

if ($Results.ServiceStatus.StoppedUnexpectedCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Stopped Services (Unexpected)</h3>"
    $html += "<table><tr><th>Service</th><th>Status</th><th>Startup Type</th><th>Note</th></tr>"
    foreach ($svc in $Results.ServiceStatus.Issues) {
        $html += "<tr class='flagged'><td><strong>$($svc.DisplayName)</strong> ($($svc.ServiceName))</td><td class='error-text'>$($svc.Status)</td><td>$($svc.StartType)</td><td class='warning-text'>$($svc.Note)</td></tr>"
    }
    $html += "</table>"
}

if ($Results.ServiceStatus.AllServices -and $Results.ServiceStatus.TotalChecked -gt 0) {
    $html += "<details><summary style='cursor:pointer;color:#3b82f6;margin-top:10px;'>View all monitored services ($($Results.ServiceStatus.TotalChecked))</summary>"
    $html += "<table><tr><th>Service</th><th>Status</th><th>Startup Type</th></tr>"
    foreach ($svc in $Results.ServiceStatus.AllServices) {
        $statusClass = if ($svc.Status -ne "Running") { "class='warning-text'" } else { "" }
        $html += "<tr><td>$($svc.DisplayName) ($($svc.ServiceName))</td><td $statusClass>$($svc.Status)</td><td>$($svc.StartType)</td></tr>"
    }
    $html += "</table></details>"
}

$html += @"
</div>

<!-- NETWORK CONFIGURATION -->
<div class="section">
    <h2><span class="icon">&#x1F310;</span> Network Configuration $(Get-StatusBadge $Results.NetworkConfig.Status)</h2>
    <p>$($Results.NetworkConfig.AdapterCount) active adapter(s). Internet: $($Results.NetworkConfig.InternetAccess). Public IP: $($Results.NetworkConfig.PublicIP).</p>
"@

if ($Results.NetworkConfig.Adapters -and $Results.NetworkConfig.AdapterCount -gt 0) {
    $html += "<table><tr><th>Adapter</th><th>IPv4</th><th>Subnet</th><th>Gateway</th><th>DNS</th><th>Link Speed</th><th>DHCP</th></tr>"
    foreach ($nic in $Results.NetworkConfig.Adapters) {
        $dhcpText = if ($nic.DHCP) { "Yes (Server: $($nic.DHCPServer))" } else { "Static" }
        $speedText = if ($nic.LinkSpeedMbps -ne "Unknown") { "$($nic.LinkSpeedMbps) Mbps" } else { "Unknown" }
        $html += "<tr><td>$($nic.Name)</td><td>$($nic.IPv4)</td><td>$($nic.SubnetMask)</td><td>$($nic.Gateway)</td><td>$($nic.DNS)</td><td>$speedText</td><td>$dhcpText</td></tr>"
    }
    $html += "</table>"

    # Show DHCP lease details for DHCP-enabled adapters
    $dhcpAdapters = $Results.NetworkConfig.Adapters | Where-Object { $_.DHCP -eq $true }
    if ($dhcpAdapters) {
        $html += "<h3 style='margin-top:12px;margin-bottom:6px;font-size:0.95em;'>DHCP Lease Details</h3><table>"
        $html += "<tr><th>Adapter</th><th>DHCP Server</th><th>Lease Obtained</th><th>Lease Expires</th></tr>"
        foreach ($nic in $dhcpAdapters) {
            $html += "<tr><td>$($nic.Name)</td><td>$($nic.DHCPServer)</td><td>$($nic.LeaseObtained)</td><td>$($nic.LeaseExpires)</td></tr>"
        }
        $html += "</table>"
    }
}

# Speed test results
if ($Results.NetworkConfig.SpeedTest.Tested) {
    $st = $Results.NetworkConfig.SpeedTest
    $dlColor = if ($st.DownloadMbps -lt 10) { "#dc3545" } elseif ($st.DownloadMbps -lt 50) { "#fd7e14" } else { "#28a745" }
    $ulColor = if ($st.UploadMbps -lt 5) { "#dc3545" } elseif ($st.UploadMbps -lt 20) { "#fd7e14" } else { "#28a745" }
    $pingColor = if ($st.PingMs -gt 100) { "#dc3545" } elseif ($st.PingMs -gt 50) { "#fd7e14" } else { "#28a745" }
    $jitterColor = if ($st.JitterMs -gt 30) { "#dc3545" } elseif ($st.JitterMs -gt 10) { "#fd7e14" } else { "#28a745" }
    $html += "<h3 style='margin-top:12px;margin-bottom:6px;font-size:0.95em;'>Internet Speed Test</h3>"
    $html += "<table>"
    $html += "<tr><td><strong>Download</strong></td><td><span style='color:$dlColor;font-weight:bold;font-size:1.1em;'>$($st.DownloadMbps) Mbps</span></td></tr>"
    $html += "<tr><td><strong>Upload</strong></td><td><span style='color:$ulColor;font-weight:bold;font-size:1.1em;'>$($st.UploadMbps) Mbps</span></td></tr>"
    $html += "<tr><td><strong>Ping (Latency)</strong></td><td><span style='color:$pingColor;font-weight:bold;'>$($st.PingMs) ms</span></td></tr>"
    $html += "<tr><td><strong>Jitter</strong></td><td><span style='color:$jitterColor;font-weight:bold;'>$($st.JitterMs) ms</span></td></tr>"
    $html += "<tr><td><strong>ISP</strong></td><td>$($st.ISP)</td></tr>"
    $html += "</table>"
} elseif ($Results.NetworkConfig.SpeedTest.Error) {
    $html += "<h3 style='margin-top:12px;margin-bottom:6px;font-size:0.95em;'>Internet Speed Test</h3>"
    $html += "<p class='detail'>Speed test could not be completed: $($Results.NetworkConfig.SpeedTest.Error)</p>"
}

# Physical printers
if ($Results.NetworkConfig.PrinterCount -gt 0) {
    $html += "<h3 style='margin-top:12px;margin-bottom:6px;font-size:0.95em;'>Physical Printers ($($Results.NetworkConfig.PrinterCount))</h3>"
    $html += "<table><tr><th>Printer</th><th>Connection</th><th>Port</th><th>Status</th><th>Default</th><th>Driver</th></tr>"
    foreach ($prt in $Results.NetworkConfig.Printers) {
        $statusColor = if ($prt.Status -match "Offline|Stopped") { "class='error-text'" } elseif ($prt.Status -eq "Idle" -or $prt.Status -eq "Printing") { "" } else { "class='warning-text'" }
        $defaultText = if ($prt.Default) { "<strong>Yes</strong>" } else { "No" }
        $html += "<tr><td><strong>$($prt.Name)</strong></td><td>$($prt.Connection)</td><td class='detail'>$($prt.Port)</td><td $statusColor>$($prt.Status)</td><td>$defaultText</td><td class='detail'>$($prt.Driver)</td></tr>"
    }
    $html += "</table>"
} else {
    $html += "<h3 style='margin-top:12px;margin-bottom:6px;font-size:0.95em;'>Physical Printers</h3>"
    $html += "<p class='detail'>No physical printers detected.</p>"
}

if ($Results.NetworkConfig.IssueCount -gt 0) {
    $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Issues Detected ($($Results.NetworkConfig.IssueCount))</h3>"
    $html += "<table><tr><th>Adapter</th><th>Issue</th><th>Severity</th></tr>"
    foreach ($issue in $Results.NetworkConfig.Issues) {
        $sevClass = if ($issue.Severity -eq "ERROR") { "class='error-text'" } else { "class='warning-text'" }
        $html += "<tr class='flagged'><td>$($issue.Adapter)</td><td>$($issue.Issue)</td><td $sevClass><strong>$($issue.Severity)</strong></td></tr>"
    }
    $html += "</table>"
}

$html += @"
</div>

<!-- ADWCLEANER -->
<div class="section">
    <h2><span class="icon">&#x1F50D;</span> AdwCleaner Scan $(Get-StatusBadge $Results.AdwCleaner.Status)</h2>
"@

if ($Results.AdwCleaner.Status -eq "SKIPPED") {
    $html += "<p class='detail'>Skipped by configuration.</p>"
} elseif ($Results.AdwCleaner.Status -eq "ERROR") {
    $html += "<p class='error-text'>Error: $($Results.AdwCleaner.Error)</p>"
} else {
    $html += "<p>Adware/Malware Detections: $($Results.AdwCleaner.DetectionCount)</p>"
    if ($Results.AdwCleaner.DetectionCount -gt 0 -and $Results.AdwCleaner.Detections) {
        $html += "<h3 style='color:#dc2626;margin:10px 0 5px;font-size:0.95em;'>Adware/Malware Found ($($Results.AdwCleaner.DetectionCount))</h3>"
        $html += "<table><tr><th>Detection</th></tr>"
        foreach ($detection in $Results.AdwCleaner.Detections) {
            $html += "<tr class='flagged'><td class='error-text'>$([System.Net.WebUtility]::HtmlEncode($detection.Trim()))</td></tr>"
        }
        $html += "</table>"
    } elseif ($Results.AdwCleaner.DetectionCount -eq 0) {
        $html += "<p style='color:#28a745;font-weight:bold;'>No adware or malware found.</p>"
    }
    if ($Results.AdwCleaner.PreinstalledCount -gt 0) {
        $html += "<h3 style='margin:10px 0 5px;font-size:0.95em;color:#6b7280;'>Preinstalled Software ($($Results.AdwCleaner.PreinstalledCount) items)</h3>"
        $html += "<table><tr><th>Item</th></tr>"
        foreach ($item in $Results.AdwCleaner.PreinstalledItems) {
            $html += "<tr><td class='detail'>$([System.Net.WebUtility]::HtmlEncode($item))</td></tr>"
        }
        $html += "</table>"
    }
    if ($Results.AdwCleaner.LogFile -ne "N/A") {
        $html += "<p class='detail'>Full log: $($Results.AdwCleaner.LogFile)</p>"
    }
}

$html += @"
</div>

<!-- RESTORE POINTS -->
<div class="section">
    <h2><span class="icon">&#x1F4BE;</span> System Restore Points $(Get-StatusBadge $Results.RestorePoints.Status)</h2>
"@

if ($Results.RestorePoints.Status -eq "ERROR") {
    $html += "<p class='error-text'>Error: $($Results.RestorePoints.Error)</p>"
} else {
    $enabledText = if ($Results.RestorePoints.Enabled) { "<span style='color:#28a745;font-weight:bold;'>Enabled</span>" } else { "<span style='color:#dc3545;font-weight:bold;'>DISABLED</span>" }
    $html += "<p>System Restore: $enabledText &bull; Restore Points: $($Results.RestorePoints.Count) &bull; Disk Usage: $($Results.RestorePoints.DiskUsage)</p>"

    if ($Results.RestorePoints.Note) {
        $noteColor = if ($Results.RestorePoints.Status -eq "WARNING") { "#dc3545" } else { "#28a745" }
        $html += "<p style='color:$noteColor;font-weight:bold;'>$($Results.RestorePoints.Note)</p>"
    }

    if ($Results.RestorePoints.Points.Count -gt 0) {
        $html += "<details><summary style='cursor:pointer;color:#0d6efd;font-weight:bold;margin:8px 0;'>Recent Restore Points ($($Results.RestorePoints.Points.Count) shown)</summary>"
        $html += "<table><tr><th>Date</th><th>Type</th><th>Description</th></tr>"
        foreach ($rp in $Results.RestorePoints.Points) {
            $html += "<tr><td>$($rp.Created)</td><td>$($rp.Type)</td><td>$([System.Net.WebUtility]::HtmlEncode($rp.Description))</td></tr>"
        }
        $html += "</table></details>"
    }
}

$html += @"
</div>

<!-- TELEMETRY SERVICES -->
<div class="section">
    <h2><span class="icon">&#x1F4E1;</span> Telemetry Services $(Get-StatusBadge $Results.TelemetryServices.Status)</h2>
"@

if ($Results.TelemetryServices.Status -eq "ERROR") {
    $html += "<p class='error-text'>Error: $($Results.TelemetryServices.Error)</p>"
} else {
    $html += "<p>Found: $($Results.TelemetryServices.TotalFound) &bull; Running: $($Results.TelemetryServices.RunningCount) &bull; Disabled: $($Results.TelemetryServices.DisabledCount)</p>"
    $html += "<p class='detail'>This is a read-only audit of Windows data-collection services. No changes have been made.</p>"

    if ($Results.TelemetryServices.Services.Count -gt 0) {
        $html += "<table><tr><th>Service</th><th>Status</th><th>Startup</th></tr>"
        foreach ($ts in $Results.TelemetryServices.Services) {
            $statusStyle = switch ($ts.Status) {
                "Running" { "color:#dc3545;font-weight:bold;" }
                "Stopped" { "color:#28a745;" }
                default   { "" }
            }
            $startStyle = if ($ts.StartType -eq "Disabled") { "color:#28a745;font-weight:bold;" } else { "" }
            $html += "<tr><td>$($ts.FriendlyName)<br/><span class='detail'>($($ts.ServiceName))</span></td><td style='$statusStyle'>$($ts.Status)</td><td style='$startStyle'>$($ts.StartType)</td></tr>"
        }
        $html += "</table>"
    }
}

$html += @"
</div>

<!-- ERRORS SUMMARY -->
"@

if ($Results.Errors.Count -gt 0) {
    $html += "<div class='section'><h2><span class='icon'>&#x26A0;</span> Errors During Run</h2><ul>"
    foreach ($err in $Results.Errors) {
        $html += "<li class='error-text'>$err</li>"
    }
    $html += "</ul></div>"
}

$html += @"

<div class="footer">
    Generated by PC Masterclass Maintenance Script v$ScriptVersion &bull; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &bull; Report: $ReportFile
</div>

</div>
</body>
</html>
"@

# Write the report
$html | Out-File -FilePath $ReportFile -Encoding UTF8 -Force

# Export results as JSON (preliminary - so it can be attached to email; re-exported after email + webhook)
$jsonFile = $ReportFile -replace '\.html$', '.json'
if ($ClientName) { $Results | Add-Member -NotePropertyName "ClientName" -NotePropertyValue $ClientName -Force }
$Results | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonFile -Encoding UTF8 -Force

# ============================================================================
# MODULE 17: EMAIL REPORT
# ============================================================================
if ($EmailTo) {
    Write-Log "Sending report via email to $EmailTo..."

    try {
        # Default EmailFrom to SmtpUser if not specified
        if (-not $EmailFrom -and $SmtpUser) { $EmailFrom = $SmtpUser }

        if (-not $SmtpUser -and -not $StoredCredential) {
            throw "SMTP credentials required. Run with -SaveCredential first, or provide -SmtpUser and -SmtpPassword."
        }

        # Build a concise plain-text summary for the email body
        $statusIcon = switch ($overallStatus) {
            "PASS"    { "[PASS]" }
            "WARNING" { "[WARNING]" }
            "FAIL"    { "[FAIL]" }
            default   { "[UNKNOWN]" }
        }

        $scriptDuration = [math]::Round(((Get-Date) - $scriptStartTime).TotalMinutes, 1)
        $clientDisplay = if ($ClientName) { $ClientName } else { $ComputerName }
        $emailSubject = "$clientDisplay - $overallStatus - Maintenance Report - $ComputerName - $(Get-Date -Format 'dd MMM yyyy') - ${scriptDuration}m - v$ScriptVersion"

        $emailBody = "[T]PC MASTERCLASS - SCHEDULED MAINTENANCE REPORT[/T]`n"
        $emailBody += "Computer".PadRight(30) + "$ComputerName ($($Results.SystemInfo.Manufacturer) $($Results.SystemInfo.Model))`n"
        $emailBody += "Serial Number".PadRight(30) + "$($Results.SystemInfo.SerialNumber)`n"
        $emailBody += "Date".PadRight(30) + "$(Get-Date -Format 'dddd, d MMMM yyyy h:mm tt')`n"
        $emailBody += "Script Version".PadRight(30) + "v$ScriptVersion`n"
        $emailBody += "OS".PadRight(30) + "$($Results.SystemInfo.OS) - Build $($Results.SystemInfo.Build)`n"
        $emailBody += "CPU".PadRight(30) + "$($Results.SystemInfo.CPU) ($($Results.SystemInfo.CPUCores)C/$($Results.SystemInfo.CPUThreads)T)`n"
        if ($Results.SystemInfo.CPUBenchmark.Score -gt 0) {
            $tierTag = switch ($Results.SystemInfo.CPUBenchmark.Tier) {
                "Very Fast"  { "[G]$($Results.SystemInfo.CPUBenchmark.Tier)[/G]" }
                "Fast"       { "[G]$($Results.SystemInfo.CPUBenchmark.Tier)[/G]" }
                "Average"    { "[W]$($Results.SystemInfo.CPUBenchmark.Tier)[/W]" }
                "Slow"       { "[W]$($Results.SystemInfo.CPUBenchmark.Tier)[/W]" }
                "Very Slow"  { "[R]$($Results.SystemInfo.CPUBenchmark.Tier)[/R]" }
                default      { $Results.SystemInfo.CPUBenchmark.Tier }
            }
            $emailBody += "CPU Benchmark".PadRight(30) + "PassMark: $($Results.SystemInfo.CPUBenchmark.Score.ToString('N0')) - $tierTag (defs: $($Results.SystemInfo.CPUBenchmark.DefinitionsAge))`n"
        }
        $emailBody += "RAM".PadRight(30) + "$($Results.SystemInfo.TotalRAM_GB) GB`n"
        $emailBody += "Uptime".PadRight(30) + "$($Results.SystemInfo.UptimeDays) days`n"
        if ($Results.SystemInfo.Temperatures[0].Zone -ne "N/A") {
            $tempItems = @()
            foreach ($t in $Results.SystemInfo.Temperatures) {
                if ($t.TempC -gt 85) {
                    $tempItems += "[R]$($t.TempC)C ($($t.Zone))[/R]"
                } elseif ($t.TempC -gt 70) {
                    $tempItems += "[W]$($t.TempC)C ($($t.Zone))[/W]"
                } else {
                    $tempItems += "$($t.TempC)C ($($t.Zone))"
                }
            }
            $emailBody += "Temperature".PadRight(30) + ($tempItems -join ", ") + "`n"
        }
        $emailBody += "`n"
                $overallColor = switch ($overallStatus) {
            "PASS"    { "[G]" }
            "WARNING" { "[W]" }
            "FAIL"    { "[R]" }
            default   { "[B]" }
        }
        $overallEnd = switch ($overallStatus) {
            "PASS"    { "[/G]" }
            "WARNING" { "[/W]" }
            "FAIL"    { "[/R]" }
            default   { "[/B]" }
        }
        $emailBody += @"

${overallColor}OVERALL STATUS: $overallStatus ($warningCount warning(s), $errorCount error(s))${overallEnd}

[H]CHECK RESULTS SUMMARY[/H]

"@
        # Build aligned two-column summary
        $checkItems = @(
            @("System File Integrity (SFC)", $Results.SFC.Status),
            @("Disk Health", $Results.DiskHealth.Status),
            @("Windows Updates", $(if ($Results.WindowsUpdates.PendingCount) { "$($Results.WindowsUpdates.Status) - $($Results.WindowsUpdates.PendingCount) pending" } else { $Results.WindowsUpdates.Status })),
            @("iDrive Backup", $Results.iDriveBackup.Status),
            @("Malwarebytes", $Results.Malwarebytes.Status),
            @("Windows Defender", $Results.Defender.Status),
            @("Event Log Errors (24h)", $Results.EventLogErrors.Status),
            @("Pending Reboot", $Results.PendingReboot.Status),
            @("Windows Firewall", $Results.Firewall.Status),
            @("User Accounts", $Results.UserAccounts.Status),
            @("DISM Component Store", $Results.DISM.Status),
            @("Temp/Cache Files", $Results.TempFiles.Status),
            @("Startup Programs", $Results.StartupPrograms.Status),
            @("Scheduled Tasks", $(if ($Results.ScheduledTasks.FlaggedCount -gt 0) { "$($Results.ScheduledTasks.Status) - $($Results.ScheduledTasks.FlaggedCount) flagged" } else { $Results.ScheduledTasks.Status })),
            @("Browser Extensions", $(if ($Results.BrowserExtensions.FlaggedCount -gt 0) { "$($Results.BrowserExtensions.Status) - $($Results.BrowserExtensions.FlaggedCount) flagged" } else { $Results.BrowserExtensions.Status })),
            @("Service Status", $(if ($Results.ServiceStatus.StoppedUnexpectedCount -gt 0) { "$($Results.ServiceStatus.Status) - $($Results.ServiceStatus.StoppedUnexpectedCount) stopped" } else { $Results.ServiceStatus.Status })),
            @("Network Config", $(if ($Results.NetworkConfig.IssueCount -gt 0) { "$($Results.NetworkConfig.Status) - $($Results.NetworkConfig.IssueCount) issue(s)" } else { $Results.NetworkConfig.Status })),
            @("AdwCleaner", $Results.AdwCleaner.Status),
            @("Restore Points", $(if ($Results.RestorePoints.Note) { "$($Results.RestorePoints.Status) - $($Results.RestorePoints.Note)" } else { $Results.RestorePoints.Status })),
            @("Telemetry Services", "$($Results.TelemetryServices.Status) - $($Results.TelemetryServices.RunningCount) running / $($Results.TelemetryServices.DisabledCount) disabled")
        )
        foreach ($item in $checkItems) {
            $label = $item[0].PadRight(30)
            $statusText = $item[1]
            # Colour-code the status in the summary
            if ($statusText -match "ERROR|FAIL") {
                $emailBody += "$label [R]$statusText[/R]`n"
            } elseif ($statusText -match "WARNING|UPDATES AVAILABLE|NOT INSTALLED|flagged|stopped") {
                $emailBody += "$label [W]$statusText[/W]`n"
            } elseif ($statusText -match "PASS|CLEAN|UP TO DATE") {
                $emailBody += "$label [G]$statusText[/G]`n"
            } elseif ($statusText -match "SKIPPED") {
                $emailBody += "$label $statusText`n"
            } else {
                $emailBody += "$label $statusText`n"
            }
        }
        $emailBody += @"

"@

        # Add iDrive backup details if installed
        if ($Results.iDriveBackup.Installed) {
            $svcStatus = if ($Results.iDriveBackup.ServiceRunning) { "[G]Running[/G]" } else { "[R]Not Running[/R]" }
            $emailBody += "`n[H]iDRIVE BACKUP DETAILS[/H]`n"
            $emailBody += "Last Backup".PadRight(30) + "$($Results.iDriveBackup.LastBackupDate)`n"
            $resultText = $Results.iDriveBackup.LastBackupResult
            if ($resultText -match "could not be backed up|failed|error") {
                $emailBody += "Result".PadRight(30) + "[W]$resultText[/W]`n"
            } else {
                $emailBody += "Result".PadRight(30) + "$resultText`n"
            }
            $emailBody += "Files Backed Up".PadRight(30) + "$($Results.iDriveBackup.FilesBackedUp) of $($Results.iDriveBackup.FilesTotal) considered`n"
            $emailBody += "Files In Sync".PadRight(30) + "$($Results.iDriveBackup.FilesInSync)`n"
            $emailBody += "Computer Name".PadRight(30) + "$($Results.iDriveBackup.ComputerName)`n"
            $emailBody += "Account".PadRight(30) + "$($Results.iDriveBackup.Account)`n"
            $emailBody += "Service".PadRight(30) + "$svcStatus`n"
            $emailBody += "`n"
        }

        # Add Malwarebytes details if installed
        if ($Results.Malwarebytes.Installed) {
            $mbSvcText = if ($Results.Malwarebytes.ServiceRunning) { "[G]Running[/G]" } else { "[R]Not Running[/R]" }
            $mbRtText = switch ($Results.Malwarebytes.RealTimeProtection) {
                "Enabled"  { "[G]Enabled[/G]" }
                "Disabled" { "[R]Disabled[/R]" }
                default    { $Results.Malwarebytes.RealTimeProtection }
            }
            $emailBody += "`n[H]MALWAREBYTES ENDPOINT PROTECTION[/H]`n"
            $emailBody += "Product".PadRight(30) + "$($Results.Malwarebytes.ProductType)`n"
            $emailBody += "Service".PadRight(30) + "$mbSvcText`n"
            $emailBody += "Real-Time Protection".PadRight(30) + "$mbRtText`n"
            $emailBody += "`n"
        }

        # Add Defender details
        if ($Results.Defender.RealTimeProtection) {
            $emailBody += "`n[H]WINDOWS DEFENDER[/H]`n"
            $rtpText = $Results.Defender.RealTimeProtection
            if ($rtpText -eq "DISABLED") {
                $emailBody += "Real-Time Protection".PadRight(30) + "[R]$rtpText[/R]`n"
            } else {
                $emailBody += "Real-Time Protection".PadRight(30) + "[G]$rtpText[/G]`n"
            }
            if ($Results.Defender.ThirdPartyAV) {
                $emailBody += "Third-Party AV".PadRight(30) + "$($Results.Defender.ThirdPartyAV)`n"
            } else {
                $emailBody += "Definitions Updated".PadRight(30) + "$($Results.Defender.DefinitionsUpdated) ($($Results.Defender.DefinitionsAge))`n"
                $emailBody += "Last Scan".PadRight(30) + "$($Results.Defender.LastScan) ($($Results.Defender.LastScanType))`n"
                $threatCount = $Results.Defender.ThreatsLast30Days
                if ($threatCount -gt 0) {
                    $emailBody += "Threats (30 days)".PadRight(30) + "[R]$threatCount[/R]`n"
                } else {
                    $emailBody += "Threats (30 days)".PadRight(30) + "[G]$threatCount[/G]`n"
                }
            }
            $emailBody += "`n"
        }

        # Add Event Log summary
        if ($Results.EventLogErrors.TotalEvents -gt 0) {
            $emailBody += "`n[H]EVENT LOG ANALYSIS (Last $($Results.EventLogErrors.HoursChecked)h)[/H]`n"
            $emailBody += "Total Events".PadRight(30) + "$($Results.EventLogErrors.TotalEvents)`n"
            if ($Results.EventLogErrors.ActionableCount -gt 0) {
                $emailBody += "Needs Attention".PadRight(30) + "[R]$($Results.EventLogErrors.ActionableCount)[/R]`n"
            } else {
                $emailBody += "Needs Attention".PadRight(30) + "[G]$($Results.EventLogErrors.ActionableCount)[/G]`n"
            }
            $emailBody += "Routine (safe to ignore)".PadRight(30) + "$($Results.EventLogErrors.RoutineCount)`n`n"

            if ($Results.EventLogErrors.ActionableCount -gt 0) {
                $emailBody += "  [R]** EVENTS NEEDING ATTENTION **[/R]`n"
                $evtCount = 0
                foreach ($evt in $Results.EventLogErrors.ActionableEvents) {
                    if ($evtCount -ge 5) { break }
                    $evtColor = if ($evt.Level -eq "Critical") { "[R]" } else { "[W]" }
                    $evtEnd = if ($evt.Level -eq "Critical") { "[/R]" } else { "[/W]" }
                    $emailBody += "  ${evtColor}[$($evt.Level)] $($evt.Source) (ID:$($evt.EventID)) at $($evt.Time)${evtEnd}`n"
                    $emailBody += "    -> $($evt.Note)`n`n"
                    $evtCount++
                }
                if ($Results.EventLogErrors.ActionableCount -gt 5) {
                    $emailBody += "  ... and $($Results.EventLogErrors.ActionableCount - 5) more (see HTML report)`n"
                }
            } else {
                $emailBody += "  All events are routine - no action required.`n"
            }
            $emailBody += "`n"
        }

        # Add Pending Reboot details
        if ($Results.PendingReboot.RebootRequired -or $Results.PendingReboot.UptimeWarning) {
            $emailBody += "`n[H]REBOOT STATUS[/H]`n"
            if ($Results.PendingReboot.RebootRequired) {
                $emailBody += "Reboot Required".PadRight(30) + "[R]YES[/R]`n"
                $emailBody += "Reasons".PadRight(30) + "[W]$($Results.PendingReboot.Reasons)[/W]`n"
            }
            if ($Results.PendingReboot.UptimeWarning) {
                $emailBody += "Uptime Warning".PadRight(30) + "[W]$($Results.PendingReboot.UptimeDays) days without reboot[/W]`n"
            }
            $emailBody += "`n"
        }

        # Add Firewall details if any profile is disabled
        if (-not $Results.Firewall.AllEnabled) {
            $emailBody += "`n[H]FIREWALL WARNING[/H]`n"
            foreach ($fwp in $Results.Firewall.Profiles) {
                if ($fwp.Enabled) {
                    $emailBody += "$($fwp.Profile)".PadRight(30) + "[G]ON[/G]`n"
                } else {
                    $emailBody += "$($fwp.Profile)".PadRight(30) + "[R]OFF *** WARNING ***[/R]`n"
                }
            }
            $emailBody += "`n"
        }

        # Add User Accounts details if flagged
        if ($Results.UserAccounts.FlaggedCount -gt 0) {
            $emailBody += "`n[H]USER ACCOUNTS WARNING[/H]`n"
            $emailBody += "Flagged Accounts".PadRight(30) + "[R]$($Results.UserAccounts.FlaggedCount)[/R]`n"
            foreach ($acct in $Results.UserAccounts.FlaggedAccounts) {
                $emailBody += "$($acct.Name)".PadRight(30) + "[W]$($acct.FlagReason) (Last logon: $($acct.LastLogon))[/W]`n"
            }
            $emailBody += "`n"
        }

        # Add DISM details if not healthy
        if ($Results.DISM.Status -ne "PASS") {
            $emailBody += "`n[H]DISM COMPONENT STORE[/H]`n"
            $emailBody += "Status".PadRight(30) + "[R]$($Results.DISM.Result)[/R]`n"
            $emailBody += "`n"
        }

        # Add Temp Files summary (always shown)
        if ($Results.TempFiles.Locations -and $Results.TempFiles.Locations.Count -gt 0) {
            $emailBody += "`n[H]TEMP AND CACHE FILES ($($Results.TempFiles.TotalSizeGB) GB total)[/H]`n"
            foreach ($loc in $Results.TempFiles.Locations) {
                if ($loc.SizeMB -gt 10) {
                    $emailBody += "$($loc.Location)".PadRight(30) + "$($loc.SizeMB) MB`n"
                }
            }
            $emailBody += "`n"
        }

        # Add BitLocker status (always shown)
        if ($Results.SystemInfo.BitLocker) {
            $emailBody += "`n[H]BITLOCKER ENCRYPTION[/H]`n"
            foreach ($blVol in $Results.SystemInfo.BitLocker) {
                $blDetail = if ($blVol.EncryptionMethod -ne "N/A") { "$($blVol.ProtectionStatus) ($($blVol.EncryptionMethod))" } else { $blVol.ProtectionStatus }
                if ($blVol.ProtectionStatus -eq "On") {
                    $emailBody += "$($blVol.Drive)".PadRight(30) + "[G]$blDetail[/G]`n"
                } elseif ($blVol.ProtectionStatus -eq "Off") {
                    $emailBody += "$($blVol.Drive)".PadRight(30) + "[W]$blDetail[/W]`n"
                } else {
                    $emailBody += "$($blVol.Drive)".PadRight(30) + "$blDetail`n"
                }
            }
            $emailBody += "`n"
        }

        # Add disk space summary (always shown)
        if ($Results.DiskHealth.Volumes) {
            $emailBody += "`n[H]DISK SPACE[/H]`n"
            foreach ($vol in $Results.DiskHealth.Volumes) {
                if ($vol.Warning) {
                    $emailBody += "  [R]$($vol.Drive) $($vol.Label) - $($vol.FreeGB)GB free of $($vol.SizeGB)GB ($($vol.FreePercent)% free) *** LOW SPACE ***[/R]`n"
                } else {
                    $emailBody += "  $($vol.Drive) ($($vol.Label)) - $($vol.FreeGB)GB free of $($vol.SizeGB)GB ($($vol.FreePercent)% free)`n"
                }
            }
            $emailBody += "`n"
        }

        # Add flagged startup items if any
        if ($Results.StartupPrograms.FlaggedCount -gt 0) {
            $emailBody += "`n[H]FLAGGED STARTUP ITEMS[/H]`n"
            foreach ($item in $Results.StartupPrograms.FlaggedItems) {
                $emailBody += "[R]$($item.Name)[/R]".PadRight(30) + "[W]$($item.Reasons)[/W]`n"
                $emailBody += "  Command: $($item.Command)`n`n"
            }
        }

        # Add Scheduled Tasks details if flagged
        if ($Results.ScheduledTasks.FlaggedCount -gt 0) {
            $emailBody += "`n[H]SCHEDULED TASKS AUDIT[/H]`n"
            $emailBody += "Third-Party Tasks".PadRight(30) + "$($Results.ScheduledTasks.ThirdPartyCount)`n"
            $emailBody += "Flagged (suspicious)".PadRight(30) + "[R]$($Results.ScheduledTasks.FlaggedCount)[/R]`n`n"
            $emailBody += "  [B]** FLAGGED TASKS **[/B]`n"
            foreach ($task in $Results.ScheduledTasks.FlaggedTasks) {
                $emailBody += "  [R]$($task.TaskName)[/R]`n"
                if ($task.Command.Length -gt 80) {
                    $emailBody += "    Command: $($task.Command.Substring(0, 80))`n"
                    $emailBody += "             $($task.Command.Substring(80))`n"
                } else {
                    $emailBody += "    Command: $($task.Command)`n"
                }
                $emailBody += "    -> [R]$($task.FlagReason)[/R]`n`n"
            }
        }

        # Add Browser Extensions details if flagged
        if ($Results.BrowserExtensions.FlaggedCount -gt 0) {
            $emailBody += "`n[H]BROWSER EXTENSION AUDIT[/H]`n"
            $emailBody += "Browsers Found".PadRight(30) + "$($Results.BrowserExtensions.BrowsersFound -join ', ')`n"
            $emailBody += "Total Extensions".PadRight(30) + "$($Results.BrowserExtensions.TotalExtensions)`n"
            $emailBody += "Flagged (review)".PadRight(30) + "[R]$($Results.BrowserExtensions.FlaggedCount)[/R]`n`n"
            foreach ($ext in $Results.BrowserExtensions.FlaggedExtensions) {
                $flagColor = if ($ext.FlagReason -match "Unknown|adware") { "[R]" } else { "[W]" }
                $flagEnd = if ($ext.FlagReason -match "Unknown|adware") { "[/R]" } else { "[/W]" }
                $emailBody += "  $($ext.Name) ($($ext.Browser))".PadRight(40) + "$flagColor$($ext.FlagReason)$flagEnd`n"
            }
            $emailBody += "`n"
        }

        # Add AdwCleaner details
        if ($Results.AdwCleaner.Status -and $Results.AdwCleaner.Status -ne "SKIPPED" -and $Results.AdwCleaner.Status -ne "ERROR") {
            $emailBody += "`n[H]ADWCLEANER SCAN[/H]`n"
            if ($Results.AdwCleaner.DetectionCount -gt 0 -and $Results.AdwCleaner.Detections) {
                $emailBody += "Adware/Malware Found".PadRight(30) + "[R]$($Results.AdwCleaner.DetectionCount)[/R]`n`n"
                foreach ($detection in $Results.AdwCleaner.Detections) {
                    $emailBody += "  [R]$($detection.Trim())[/R]`n"
                }
                $emailBody += "`n"
            } else {
                $emailBody += "[G]No adware or malware found.[/G]`n`n"
            }
            if ($Results.AdwCleaner.PreinstalledCount -gt 0) {
                $emailBody += "Preinstalled Software".PadRight(30) + "$($Results.AdwCleaner.PreinstalledCount) item(s)`n"
                foreach ($item in $Results.AdwCleaner.PreinstalledItems) {
                    $emailBody += "  $item`n"
                }
                $emailBody += "`n"
            }
        }

        # Add Service Status details if issues found
        if ($Results.ServiceStatus.StoppedUnexpectedCount -gt 0) {
            $emailBody += "`n[H]SERVICE STATUS MONITOR[/H]`n"
            $emailBody += "Total Services Checked".PadRight(30) + "$($Results.ServiceStatus.TotalChecked)`n"
            $emailBody += "Running".PadRight(30) + "$($Results.ServiceStatus.RunningCount)`n"
            $emailBody += "Stopped (unexpected)".PadRight(30) + "[R]$($Results.ServiceStatus.StoppedUnexpectedCount)[/R]`n`n"
            $emailBody += "  [B]** ISSUES **[/B]`n"
            foreach ($svc in $Results.ServiceStatus.Issues) {
                $emailBody += "  [R]$($svc.DisplayName) ($($svc.ServiceName))[/R]`n"
                $emailBody += "    Status: $($svc.Status), Startup: $($svc.StartType)`n"
                $emailBody += "    -> [R]$($svc.Note)[/R]`n`n"
            }
        }

        # Add Network Config details (always shown for visibility)
        if ($Results.NetworkConfig.AdapterCount -gt 0) {
            $emailBody += "`n[H]NETWORK CONFIGURATION[/H]`n"
            $emailBody += "Active Adapters".PadRight(30) + "$($Results.NetworkConfig.AdapterCount)`n"
            $emailBody += "Internet Access".PadRight(30) + "$($Results.NetworkConfig.InternetAccess)`n"
            $emailBody += "DNS Reachable".PadRight(30) + "$($Results.NetworkConfig.DNSReachable)`n"
            $emailBody += "Public IP".PadRight(30) + "$($Results.NetworkConfig.PublicIP)`n`n"
            foreach ($nic in $Results.NetworkConfig.Adapters) {
                $speedText = if ($nic.LinkSpeedMbps -ne "Unknown") { "$($nic.LinkSpeedMbps) Mbps" } else { "Unknown" }
                $dhcpText = if ($nic.DHCP) { "DHCP" } else { "Static" }
                $emailBody += "  [B]$($nic.Name)[/B]`n"
                $emailBody += "    IPv4: $($nic.IPv4) / $($nic.SubnetMask) | Gateway: $($nic.Gateway)`n"
                $emailBody += "    DNS: $($nic.DNS) | Link: $speedText | $dhcpText`n"
                if ($nic.DHCP -and $nic.LeaseExpires -ne "N/A") {
                    $emailBody += "    DHCP Lease: $($nic.LeaseObtained) to $($nic.LeaseExpires)`n"
                }
                $emailBody += "`n"
            }
            # Speed test results
            if ($Results.NetworkConfig.SpeedTest.Tested) {
                $st = $Results.NetworkConfig.SpeedTest
                $dlText = if ($st.DownloadMbps -lt 10) { "[R]$($st.DownloadMbps) Mbps[/R]" } elseif ($st.DownloadMbps -lt 50) { "[W]$($st.DownloadMbps) Mbps[/W]" } else { "[G]$($st.DownloadMbps) Mbps[/G]" }
                $ulText = if ($st.UploadMbps -lt 5) { "[R]$($st.UploadMbps) Mbps[/R]" } elseif ($st.UploadMbps -lt 20) { "[W]$($st.UploadMbps) Mbps[/W]" } else { "[G]$($st.UploadMbps) Mbps[/G]" }
                $pingText = if ($st.PingMs -gt 100) { "[R]$($st.PingMs) ms[/R]" } elseif ($st.PingMs -gt 50) { "[W]$($st.PingMs) ms[/W]" } else { "[G]$($st.PingMs) ms[/G]" }
                $jitterText = if ($st.JitterMs -gt 30) { "[R]$($st.JitterMs) ms[/R]" } elseif ($st.JitterMs -gt 10) { "[W]$($st.JitterMs) ms[/W]" } else { "[G]$($st.JitterMs) ms[/G]" }
                $emailBody += "  [B]Internet Speed Test[/B]`n"
                $emailBody += "  Download".PadRight(30) + "$dlText`n"
                $emailBody += "  Upload".PadRight(30) + "$ulText`n"
                $emailBody += "  Ping (Latency)".PadRight(30) + "$pingText`n"
                $emailBody += "  Jitter".PadRight(30) + "$jitterText`n"
                $emailBody += "  ISP".PadRight(30) + "$($st.ISP)`n`n"
            }

            # Physical printers
            if ($Results.NetworkConfig.PrinterCount -gt 0) {
                $emailBody += "  [B]Physical Printers ($($Results.NetworkConfig.PrinterCount))[/B]`n"
                foreach ($prt in $Results.NetworkConfig.Printers) {
                    $defaultTag = if ($prt.Default) { " [default]" } else { "" }
                    if ($prt.Status -match "Offline|Stopped") {
                        $emailBody += "  [R]$($prt.Name)[/R]".PadRight(36) + "$($prt.Connection), [R]$($prt.Status)[/R]$defaultTag`n"
                    } else {
                        $emailBody += "  $($prt.Name)".PadRight(30) + "$($prt.Connection), $($prt.Status)$defaultTag`n"
                    }
                }
                $emailBody += "`n"
            } else {
                $emailBody += "  [B]Physical Printers[/B]        None detected`n`n"
            }

            if ($Results.NetworkConfig.IssueCount -gt 0) {
                $emailBody += "  [R]** NETWORK ISSUES DETECTED **[/R]`n"
                foreach ($issue in $Results.NetworkConfig.Issues) {
                    $issueColor = if ($issue.Severity -eq "ERROR") { "[R]" } else { "[W]" }
                    $issueEnd = if ($issue.Severity -eq "ERROR") { "[/R]" } else { "[/W]" }
                    $emailBody += "  ${issueColor}[$($issue.Severity)] $($issue.Adapter): $($issue.Issue)${issueEnd}`n"
                }
                $emailBody += "`n"
            }
        }

        # Add Restore Point details
        if ($Results.RestorePoints.Status -and $Results.RestorePoints.Status -ne "ERROR") {
            $emailBody += "`n[H]SYSTEM RESTORE POINTS[/H]`n"
            $srText = if ($Results.RestorePoints.Enabled) { "[G]Enabled[/G]" } else { "[R]DISABLED[/R]" }
            $emailBody += "System Restore".PadRight(30) + "$srText`n"
            $emailBody += "Restore Points".PadRight(30) + "$($Results.RestorePoints.Count)`n"
            $emailBody += "Disk Usage".PadRight(30) + "$($Results.RestorePoints.DiskUsage)`n"
            if ($Results.RestorePoints.Note) {
                $noteTag = if ($Results.RestorePoints.Status -eq "WARNING") { "[W]$($Results.RestorePoints.Note)[/W]" } else { "[G]$($Results.RestorePoints.Note)[/G]" }
                $emailBody += "Note".PadRight(30) + "$noteTag`n"
            }
            if ($Results.RestorePoints.Points.Count -gt 0) {
                $emailBody += "`n  Recent restore points:`n"
                foreach ($rp in $Results.RestorePoints.Points) {
                    $emailBody += "  $($rp.Created)".PadRight(28) + "$($rp.Type) - $($rp.Description)`n"
                }
            }
            $emailBody += "`n"
        }

        # Add Telemetry Services details
        if ($Results.TelemetryServices.Status -and $Results.TelemetryServices.Status -ne "ERROR") {
            $emailBody += "`n[H]TELEMETRY SERVICES (read-only audit)[/H]`n"
            $emailBody += "Services Found".PadRight(30) + "$($Results.TelemetryServices.TotalFound)`n"
            $emailBody += "Running".PadRight(30) + "$($Results.TelemetryServices.RunningCount)`n"
            $emailBody += "Disabled".PadRight(30) + "$($Results.TelemetryServices.DisabledCount)`n`n"
            foreach ($ts in $Results.TelemetryServices.Services) {
                $svcStatus = switch ($ts.Status) {
                    "Running" { "[W]Running[/W]" }
                    "Stopped" { "[G]Stopped[/G]" }
                    default   { $ts.Status }
                }
                $emailBody += "  $($ts.FriendlyName)`n"
                $emailBody += "    ($($ts.ServiceName))".PadRight(30) + "$svcStatus | Startup: $($ts.StartType)`n"
            }
            $emailBody += "`n"
        }

        # Add errors if any
        if ($Results.Errors.Count -gt 0) {
            $emailBody += "`n[R]ERRORS DURING RUN[/R]`n"
            foreach ($err in $Results.Errors) {
                $emailBody += "  [R]- $err[/R]`n"
            }
            $emailBody += "`n"
        }

        $emailBody += @"

Full HTML report and JSON data are attached.
Local report path: $ReportFile

Generated by PC Masterclass Maintenance Script v$ScriptVersion
"@

        # Build credentials  - use stored credential if available, otherwise from parameters
        $credential = $null
        if ($StoredCredential -and $StoredCredential.Credential) {
            $credential = $StoredCredential.Credential
            Write-Log "Using stored encrypted credentials for $SmtpUser"
        } elseif ($SmtpPassword) {
            $securePass = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($SmtpUser, $securePass)
        } else {
            throw "No credentials available. Run with -SaveCredential or provide -SmtpPassword."
        }

        # Collect attachments  - HTML report and JSON
        $attachments = @($ReportFile)
        if (Test-Path $jsonFile) { $attachments += $jsonFile }

        # Wrap email body in HTML with monospaced font for column alignment
        # Escape any HTML-special characters in the data, then restore our bold markers
        $safeBody = [System.Net.WebUtility]::HtmlEncode($emailBody)
        $safeBody = $safeBody -replace '\[B\]', '<b>' -replace '\[/B\]', '</b>'
        $safeBody = $safeBody -replace '\[H\]', '<b style="font-size:19px;">' -replace '\[/H\]', '</b>'
        $safeBody = $safeBody -replace '\[T\]', '<b style="font-size:23px;">' -replace '\[/T\]', '</b>'
        $safeBody = $safeBody -replace '\[R\]', '<b style="color:#dc3545;">' -replace '\[/R\]', '</b>'
        $safeBody = $safeBody -replace '\[W\]', '<b style="color:#fd7e14;">' -replace '\[/W\]', '</b>'
        $safeBody = $safeBody -replace '\[G\]', '<b style="color:#28a745;">' -replace '\[/G\]', '</b>'
        # Build machine-readable status block for dashboard parsing
        $diskFreeGB = 0
        $diskTotalGB = 0
        if ($Results.DiskHealth -and $Results.DiskHealth.Volumes) {
            $sysDrive = $Results.DiskHealth.Volumes | Where-Object { $_.Drive -eq "C:" } | Select-Object -First 1
            if ($sysDrive) {
                $diskFreeGB = [math]::Round($sysDrive.FreeGB, 1)
                $diskTotalGB = [math]::Round($sysDrive.SizeGB, 1)
            }
        }
        $dashboardStatus = @{
            v          = 1
            computer   = $ComputerName
            client     = if ($ClientName) { $ClientName } else { $ComputerName }
            status     = $overallStatus
            version    = $ScriptVersion
            timestamp  = (Get-Date -Format "o")
            duration   = $scriptDuration
            warnings   = $warningCount
            errors     = $errorCount
            os         = "$($Results.SystemInfo.OS)"
            cpu        = "$($Results.SystemInfo.CPU)"
            cpuScore   = if ($Results.SystemInfo.CPUBenchmark) { $Results.SystemInfo.CPUBenchmark.Score } else { 0 }
            cpuTier    = if ($Results.SystemInfo.CPUBenchmark) { $Results.SystemInfo.CPUBenchmark.Tier } else { "Unknown" }
            ramGB      = $Results.SystemInfo.TotalRAM_GB
            diskFreeGB = $diskFreeGB
            diskTotalGB = $diskTotalGB
            uptimeDays = $Results.SystemInfo.UptimeDays
        } | ConvertTo-Json -Compress
        $statusComment = "<!--PCMC_DASHBOARD:${dashboardStatus}:PCMC_DASHBOARD-->"

        $htmlEmailBody = @"
<html><body>
<pre style="font-family: Consolas, 'Courier New', monospace; font-size: 15px; line-height: 1.5; color: #333;">
$safeBody
</pre>
$statusComment
</body></html>
"@

        # Send the email
        $mailParams = @{
            From        = $EmailFrom
            To          = $EmailTo
            Subject     = $emailSubject
            Body        = $htmlEmailBody
            BodyAsHtml  = $true
            SmtpServer  = $SmtpServer
            Port        = $SmtpPort
            UseSsl      = $true
            Credential  = $credential
            Attachments = $attachments
        }

        Send-MailMessage @mailParams -ErrorAction Stop

        Write-Log "Email sent successfully to $EmailTo"
        $Results.EmailResult = @{
            Status   = "Sent"
            To       = $EmailTo
            From     = $EmailFrom
            Subject  = $emailSubject
            SmtpServer = $SmtpServer
        }

    } catch {
        Write-Log "Failed to send email: $_" "ERROR"
        $Results.Errors += "Email: $_"
        $Results.EmailResult = @{
            Status = "FAILED"
            Error  = $_.ToString()
            To     = $EmailTo
        }
    }
} else {
    Write-Log "Email not configured (no -EmailTo specified). Skipping."
    $Results.EmailResult = @{ Status = "Not configured" }
}


# ============================================================================
# UPDATE ROLLOUT TRACKER (webhook to Google Sheets)
# ============================================================================
$WebhookUrl = "https://script.google.com/macros/s/AKfycbyKkPyodUa3M2Ka9vXzgjMK0hzq6EfA58unifA7Ih6h4OxjLYfXuqea8rrcO2i4yMmF/exec"
$WebhookSecret = "pcm-tracker-2026"

try {
    $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    $webhookData = @{
        secret         = $WebhookSecret
        computerName   = $env:COMPUTERNAME
        scriptVersion  = $ScriptVersion
        lastRun        = (Get-Date).ToString("yyyy-MM-dd")
        frequencyDays  = 90
        osVersion      = if ($osCaption) { $osCaption } else { "Unknown" }
    } | ConvertTo-Json -Compress

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $webhookData -ContentType "application/json" -TimeoutSec 30 -ErrorAction Stop

    if ($response.status -eq "ok") {
        Write-Log "Rollout Tracker updated successfully (row $($response.row))"
        $Results.WebhookResult = @{
            Status   = "OK"
            Row      = $response.row
            Response = $response.message
            URL      = $WebhookUrl
        }
    } else {
        Write-Log "Rollout Tracker update warning: $($response.message)" "WARN"
        $Results.WebhookResult = @{
            Status   = "WARNING"
            Response = $response.message
            URL      = $WebhookUrl
        }
    }
} catch {
    Write-Log "Rollout Tracker webhook failed (non-critical): $_" "WARN"
    $Results.WebhookResult = @{
        Status = "FAILED"
        Error  = $_.ToString()
        URL    = $WebhookUrl
    }
}

# Export results as JSON (after email + webhook so all results are captured)
$Results | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonFile -Encoding UTF8 -Force

Write-Log "============================================"
Write-Log "REPORT SAVED: $ReportFile"
Write-Log "JSON DATA:    $jsonFile"
Write-Log "OVERALL:      $overallStatus ($warningCount warnings, $errorCount errors)"
if ($EmailTo) { Write-Log "EMAILED TO:   $EmailTo" }
Write-Log "============================================"

# Return the report path for automation
return $ReportFile
