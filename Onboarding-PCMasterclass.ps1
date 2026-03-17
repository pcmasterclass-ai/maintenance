<#
.SYNOPSIS
    PC Masterclass - Client Onboarding / Deployment Script
.DESCRIPTION
    One-step deployment script for new clients. Run this once on a client machine
    (via TeamViewer or in person) to set up the full PC Masterclass maintenance
    environment. It creates the folder structure, downloads the latest maintenance
    script from GitHub, and confirms everything is ready.

    After deployment, the maintenance script can be run at any time with:
      powershell -ExecutionPolicy Bypass -File "C:\Teamviewer\PCMasterclass-Maintenance.ps1"

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.3.5
    Date:    2026-03-16

    USAGE (paste into an elevated PowerShell prompt):
      powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\Onboarding-PCMasterclass.ps1"

    WHAT THIS SCRIPT DOES:
      1. Creates C:\Teamviewer folder structure (Reports, Config, Tools)
      2. Downloads the latest PCMasterclass-Maintenance.ps1 from GitHub
      3. Verifies the download and displays the installed version
      4. Runs a quick environment check (PowerShell version, admin rights, disk space)
      5. Displays machine info for the Rollout Tracker
      6. Sets up email credentials (DPAPI + AES fallback for SYSTEM)
      7. Sends a test email to verify SMTP is working BEFORE proceeding
      8. Creates a Windows Scheduled Task for unattended maintenance runs

    WHAT THIS SCRIPT DOES NOT DO:
      - It does not run the maintenance script (you decide when to run it)
      - It does not modify system settings beyond the scheduled task
#>

# ============================================================================
# CONFIGURATION
# ============================================================================
$DeployVersion = "1.3.5"
$BaseDir = "C:\Teamviewer"
$ScriptName = "PCMasterclass-Maintenance.ps1"
$GitHubRepo = "pcmasterclass-ai/maintenance"
$GitHubBranch = "main"
$GitHubToken = ""  # Not needed - repo is public

# Email & scheduling defaults
$DefaultEmailTo = "reports@pcmasterclass.com.au"
$DefaultRunTime = "1:00AM"
$DefaultFrequencyDays = 90  # Quarterly (every 90 days)
$TaskName = "PCMasterclass-Maintenance"

# Derived paths
$ScriptPath = Join-Path $BaseDir $ScriptName
$ReportsDir = Join-Path $BaseDir "Reports"
$ConfigDir  = Join-Path $BaseDir "Config"
$ToolsDir   = Join-Path $BaseDir "Tools"
$DataDir    = Join-Path $BaseDir "Data"
$DownloadUrl = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch/$ScriptName"
$BenchmarkCsvUrl = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch/cpu-benchmarks.csv"
$BenchmarkCsvPath = Join-Path $DataDir "cpu-benchmarks.csv"

# ============================================================================
# DISPLAY BANNER
# ============================================================================
function Show-Banner {
    $banner = @"

    ====================================================
       PC Masterclass - Client Deployment Script v$DeployVersion
    ====================================================
       Setting up automated maintenance environment...
    ====================================================

"@
    Write-Host $banner -ForegroundColor Cyan
}

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
# PRE-FLIGHT CHECKS
# ============================================================================
function Test-Prerequisites {
    Write-Section "PRE-FLIGHT CHECKS"
    $allGood = $true

    # Check admin rights
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
    if ($isAdmin) {
        Write-OK "Running as Administrator"
    } else {
        Write-Fail "Not running as Administrator - this script requires elevation"
        $allGood = $false
    }

    # Check PowerShell version
    $psVer = $PSVersionTable.PSVersion
    if ($psVer.Major -ge 5) {
        Write-OK "PowerShell version: $psVer"
    } else {
        Write-Warn "PowerShell version $psVer detected (5.1+ recommended)"
    }

    # Check Windows version
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($osInfo) {
        Write-OK "OS: $($osInfo.Caption) (Build $($osInfo.BuildNumber))"
    }

    # Check internet connectivity
    Write-Step "Testing internet connectivity..."
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $testUrl = "https://raw.githubusercontent.com"
        $request = [System.Net.WebRequest]::Create($testUrl)
        $request.Timeout = 10000
        $response = $request.GetResponse()
        $response.Close()
        Write-OK "Internet connectivity confirmed"
    } catch {
        Write-Fail "Cannot reach GitHub - check internet connection"
        $allGood = $false
    }

    # Check disk space on C:
    $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    if ($disk) {
        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
        if ($freeGB -gt 5) {
            Write-OK "Disk space: ${freeGB}GB free on C:"
        } else {
            Write-Warn "Low disk space: ${freeGB}GB free on C: (recommend 5GB+)"
        }
    }

    return $allGood
}

# ============================================================================
# CREATE FOLDER STRUCTURE
# ============================================================================
function New-FolderStructure {
    Write-Section "CREATING FOLDER STRUCTURE"

    $folders = @($BaseDir, $ReportsDir, $ConfigDir, $ToolsDir, $DataDir)

    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            Write-OK "Already exists: $folder"
        } else {
            try {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
                Write-OK "Created: $folder"
            } catch {
                Write-Fail "Failed to create: $folder - $($_.Exception.Message)"
                return $false
            }
        }
    }

    return $true
}

# ============================================================================
# DOWNLOAD MAINTENANCE SCRIPT
# ============================================================================
function Get-MaintenanceScript {
    Write-Section "DOWNLOADING MAINTENANCE SCRIPT"

    # Check if script already exists
    $existingVersion = $null
    if (Test-Path $ScriptPath) {
        $existingContent = Get-Content $ScriptPath -Raw -ErrorAction SilentlyContinue
        if ($existingContent -match '\$ScriptVersion\s*=\s*"([^"]+)"') {
            $existingVersion = $Matches[1]
        }
        Write-Step "Existing installation found (v$existingVersion) - will update if newer available"
    } else {
        Write-Step "Fresh install - downloading latest version..."
    }

    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

        $webClient = New-Object System.Net.WebClient
        if ($GitHubToken) {
            $webClient.Headers.Add("Authorization", "token $GitHubToken")
        }

        # Download to temp first, then move (safer)
        $tempFile = Join-Path $env:TEMP "PCMasterclass-Maintenance-deploy.ps1"
        $webClient.DownloadFile($DownloadUrl, $tempFile)

        # Verify download
        if (-not (Test-Path $tempFile)) {
            Write-Fail "Download failed - temp file not created"
            return $false
        }

        $fileSize = (Get-Item $tempFile).Length
        if ($fileSize -lt 1000) {
            Write-Fail "Downloaded file is too small ($fileSize bytes) - may be an error page"
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $false
        }

        # Extract version from downloaded script
        $downloadedContent = Get-Content $tempFile -Raw
        $downloadedVersion = "unknown"
        if ($downloadedContent -match '\$ScriptVersion\s*=\s*"([^"]+)"') {
            $downloadedVersion = $Matches[1]
        }

        # Copy to final location
        Copy-Item -Path $tempFile -Destination $ScriptPath -Force
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

        # Remove Zone.Identifier so RemoteSigned execution policy won't block the script
        Unblock-File -Path $ScriptPath -ErrorAction SilentlyContinue

        if ($existingVersion -and $existingVersion -eq $downloadedVersion) {
            Write-OK "Script is already up to date (v$downloadedVersion)"
        } elseif ($existingVersion) {
            Write-OK "Updated: v$existingVersion -> v$downloadedVersion"
        } else {
            Write-OK "Installed: PCMasterclass-Maintenance.ps1 v$downloadedVersion"
        }

        Write-OK "Location: $ScriptPath"
        Write-OK "File size: $([math]::Round($fileSize / 1KB, 1))KB"

        # Download CPU benchmarks CSV (used for CPU performance tier rating)
        Write-Step "Downloading CPU benchmark definitions..."
        try {
            $csvClient = New-Object System.Net.WebClient
            if ($GitHubToken) { $csvClient.Headers.Add("Authorization", "token $GitHubToken") }
            $csvClient.DownloadFile($BenchmarkCsvUrl, $BenchmarkCsvPath)
            Unblock-File -Path $BenchmarkCsvPath -ErrorAction SilentlyContinue
            $csvSize = (Get-Item $BenchmarkCsvPath).Length
            $csvLines = (Get-Content $BenchmarkCsvPath | Measure-Object).Count - 1
            Write-OK "CPU benchmarks: $csvLines processors ($([math]::Round($csvSize / 1KB, 1))KB)"
        } catch {
            Write-Warn "Could not download CPU benchmarks: $($_.Exception.Message)"
            Write-Warn "The maintenance script will attempt to download it on first run"
        }

        return $true

    } catch {
        Write-Fail "Download failed: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# POST-INSTALL VERIFICATION
# ============================================================================
function Test-Installation {
    Write-Section "VERIFYING INSTALLATION"
    $allGood = $true

    # Check script exists
    if (Test-Path $ScriptPath) {
        Write-OK "Maintenance script present"
    } else {
        Write-Fail "Maintenance script missing from $ScriptPath"
        $allGood = $false
    }

    # Check all folders exist
    foreach ($dir in @($BaseDir, $ReportsDir, $ConfigDir, $ToolsDir)) {
        if (Test-Path $dir) {
            Write-OK "Directory OK: $dir"
        } else {
            Write-Fail "Directory missing: $dir"
            $allGood = $false
        }
    }

    # Check script is parseable (not corrupted)
    if (Test-Path $ScriptPath) {
        try {
            $null = [System.Management.Automation.PSParser]::Tokenize(
                (Get-Content $ScriptPath -Raw), [ref]$null
            )
            Write-OK "Script syntax validation passed"
        } catch {
            Write-Fail "Script may be corrupted - syntax check failed"
            $allGood = $false
        }
    }

    return $allGood
}

# ============================================================================
# GATHER MACHINE INFO (for Rollout Tracker)
# ============================================================================
function Get-MachineInfo {
    Write-Section "MACHINE INFORMATION"

    $computerName = $env:COMPUTERNAME
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $csInfo = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue

    $osVersion = if ($osInfo) { "$($osInfo.Caption) ($($osInfo.Version))" } else { "Unknown" }
    $deviceType = if ($csInfo) {
        switch ($csInfo.PCSystemType) {
            1 { "Desktop" }
            2 { "Laptop" }
            3 { "Workstation" }
            default { "Unknown ($($csInfo.PCSystemType))" }
        }
    } else { "Unknown" }
    $manufacturer = if ($csInfo) { "$($csInfo.Manufacturer) $($csInfo.Model)" } else { "Unknown" }

    Write-OK "Computer Name: $computerName"
    Write-OK "Device Type:   $deviceType"
    Write-OK "Hardware:      $manufacturer"
    Write-OK "OS Version:    $osVersion"

    Write-Host ""
    Write-Host "  Copy these details to your Rollout Tracker:" -ForegroundColor Yellow
    Write-Host "    Computer Name:  $computerName" -ForegroundColor White
    Write-Host "    Device:         $deviceType" -ForegroundColor White
    Write-Host "    OS Version:     $osVersion" -ForegroundColor White
}

# ============================================================================
# EMAIL CREDENTIAL SETUP
# ============================================================================
function Set-EmailCredentials {
    Write-Section "EMAIL CREDENTIAL SETUP"
    Write-Host ""
    Write-Host "  The maintenance script can email reports to you after each run." -ForegroundColor White
    Write-Host "  Credentials are encrypted and stored locally on this machine only." -ForegroundColor White
    Write-Host ""

    $setupEmail = Read-Host "  Set up email credentials now? (Y/N)"

    if ($setupEmail -notmatch '^[Yy]') {
        Write-Warn "Skipped - email will not work until credentials are configured"
        $script:emailConfigured = $false
        return
    }

    Write-Host ""
    Write-Host "  Enter the SMTP credentials for sending reports." -ForegroundColor White
    Write-Host "  For Gmail, use an App Password (not your regular password)." -ForegroundColor White
    Write-Host ""
    Write-Host "  IMPORTANT: Enter the App Password WITHOUT SPACES." -ForegroundColor Yellow
    Write-Host "  Google displays it as 'abcd efgh ijkl mnop' but enter: abcdefghijklmnop" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  NOTE: Use the actual Google account (paul@), NOT the alias (reports@)." -ForegroundColor Yellow
    Write-Host "  Gmail requires the real account for SMTP authentication." -ForegroundColor Yellow
    Write-Host ""

    $defaultEmail = "paul@pcmasterclass.com.au"
    $smtpInput = Read-Host "  SMTP login account (press Enter for $defaultEmail)"
    $smtpUser = if ($smtpInput) { $smtpInput } else { $defaultEmail }
    Write-OK "Using: $smtpUser"

    $smtpPassword = Read-Host "  App Password (no spaces)" -AsSecureString
    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($smtpPassword)
    )
    if (-not $plainPassword) {
        Write-Warn "No password entered - skipping credential setup"
        $script:emailConfigured = $false
        return
    }

    # Strip any spaces the user may have included
    $plainPassword = $plainPassword -replace '\s', ''

    Write-Step "Saving encrypted credentials..."

    try {
        $secPass = ConvertTo-SecureString $plainPassword -AsPlainText -Force

        # Save DPAPI credential (for interactive runs under this user account)
        Write-Step "Saving DPAPI credentials (interactive use)..."
        $dpapiCred = @{
            SmtpUser   = $smtpUser
            SmtpServer = "smtp.gmail.com"
            SmtpPort   = 587
            EmailFrom  = "reports@pcmasterclass.com.au"
            Credential = New-Object System.Management.Automation.PSCredential(
                $smtpUser,
                $secPass
            )
        }
        $dpapiCred | Export-Clixml -Path (Join-Path $ConfigDir "smtp-cred.xml") -Force
        Write-OK "DPAPI credentials saved (for interactive runs)"

        # Save AES-encrypted fallback for SYSTEM scheduled task
        # DPAPI credentials are tied to the user account - SYSTEM can't read them
        Write-Step "Saving AES fallback credentials (for SYSTEM scheduled task)..."
        $aesKeyPath    = Join-Path $ConfigDir "smtp.key"
        $aesConfigPath = Join-Path $ConfigDir "smtp-config.xml"

        # Generate a random 256-bit AES key
        $aesKey = New-Object byte[] 32
        [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($aesKey)

        # Encrypt the password with the AES key (not tied to any user account)
        $encPass = ConvertFrom-SecureString $secPass -Key $aesKey

        # Save the AES key file
        Set-Content -Path $aesKeyPath -Value $aesKey -Encoding Byte -Force

        # Save the config with encrypted password
        $aesConfig = @{
            SmtpUser      = $smtpUser
            SmtpServer    = "smtp.gmail.com"
            SmtpPort      = 587
            EmailFrom     = "reports@pcmasterclass.com.au"
            EncryptedPass = $encPass
        }
        $aesConfig | Export-Clixml -Path $aesConfigPath -Force

        # Restrict permissions on credential files to SYSTEM and current user only
        try {
            foreach ($file in @($aesKeyPath, $aesConfigPath, (Join-Path $ConfigDir "smtp-cred.xml"))) {
                if (Test-Path $file) {
                    $acl = Get-Acl $file
                    $acl.SetAccessRuleProtection($true, $false)
                    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    $userRule   = New-Object System.Security.AccessControl.FileSystemAccessRule($currentUser, "FullControl", "None", "None", "Allow")
                    $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "None", "None", "Allow")
                    $acl.AddAccessRule($userRule)
                    $acl.AddAccessRule($systemRule)
                    Set-Acl $file $acl
                }
            }
        } catch {
            Write-Warn "Could not restrict credential file permissions: $_"
        }

        Write-OK "AES fallback credentials saved (for SYSTEM scheduled task)"

        # Verify all three credential files exist
        $credFiles = @(
            (Join-Path $ConfigDir "smtp-cred.xml"),
            (Join-Path $ConfigDir "smtp.key"),
            (Join-Path $ConfigDir "smtp-config.xml")
        )
        $allExist = $true
        foreach ($f in $credFiles) {
            if (-not (Test-Path $f)) {
                Write-Fail "Missing: $f"
                $allExist = $false
            }
        }
        if ($allExist) {
            Write-OK "All credential files verified"
        }

        # Store for use by email test
        $script:smtpUser = $smtpUser
        $script:plainPassword = $plainPassword
        $script:emailConfigured = $true

    } catch {
        Write-Fail "Failed to save credentials: $($_.Exception.Message)"
        Write-Warn "You can set this up later manually"
        $script:emailConfigured = $false
    }
}

# ============================================================================
# EMAIL SEND TEST
# ============================================================================
function Test-EmailSend {
    Write-Section "EMAIL SEND TEST"

    if (-not $script:emailConfigured) {
        Write-Warn "Email credentials not configured - skipping test"
        return $false
    }

    Write-Step "Sending test email to $DefaultEmailTo..."

    try {
        $secPass = ConvertTo-SecureString $script:plainPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential(
            $script:smtpUser,
            $secPass
        )

        $computerName = $env:COMPUTERNAME
        Send-MailMessage `
            -From $script:smtpUser `
            -To $DefaultEmailTo `
            -Subject "PC Masterclass - Deployment Test - $computerName" `
            -Body "Email test successful for $computerName at $(Get-Date). This machine is ready to send maintenance reports." `
            -SmtpServer "smtp.gmail.com" `
            -Port 587 `
            -UseSsl `
            -Credential $credential `
            -ErrorAction Stop

        Write-OK "Test email sent successfully to $DefaultEmailTo"
        Write-OK "Check your inbox to confirm delivery"
        return $true

    } catch {
        Write-Fail "Email test FAILED: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "  Common causes:" -ForegroundColor Yellow
        Write-Host "    - App Password entered with spaces (must be 16 chars, no spaces)" -ForegroundColor White
        Write-Host "    - App Password generated for wrong Google account" -ForegroundColor White
        Write-Host "    - 2-Step Verification not enabled on the Google account" -ForegroundColor White
        Write-Host "    - No internet / firewall blocking port 587" -ForegroundColor White
        Write-Host ""

        $retry = Read-Host "  Re-enter App Password and try again? (Y/N)"
        if ($retry -match '^[Yy]') {
            Write-Host ""
            Write-Host "  REMINDER: Enter WITHOUT spaces (e.g. abcdefghijklmnop)" -ForegroundColor Yellow
            $retryPassword = Read-Host "  App Password (no spaces)" -AsSecureString
            $retryPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($retryPassword)
            )
            $retryPlain = $retryPlain -replace '\s', ''

            if ($retryPlain) {
                # Update stored credentials with new password
                $script:plainPassword = $retryPlain
                $secPass2 = ConvertTo-SecureString $retryPlain -AsPlainText -Force

                # Re-save DPAPI credential
                @{
                    SmtpUser   = $script:smtpUser
                    SmtpServer = "smtp.gmail.com"
                    SmtpPort   = 587
                    EmailFrom  = "reports@pcmasterclass.com.au"
                    Credential = New-Object System.Management.Automation.PSCredential(
                        $script:smtpUser,
                        $secPass2
                    )
                } | Export-Clixml -Path (Join-Path $ConfigDir "smtp-cred.xml") -Force

                # Re-save AES credential
                $aesKeyPath = Join-Path $ConfigDir "smtp.key"
                $aesKey = New-Object byte[] 32
                [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($aesKey)
                Set-Content -Path $aesKeyPath -Value $aesKey -Encoding Byte -Force
                $enc2 = ConvertFrom-SecureString $secPass2 -Key $aesKey
                @{
                    SmtpUser      = $script:smtpUser
                    SmtpServer    = "smtp.gmail.com"
                    SmtpPort      = 587
                    EmailFrom     = "reports@pcmasterclass.com.au"
                    EncryptedPass = $enc2
                } | Export-Clixml -Path (Join-Path $ConfigDir "smtp-config.xml") -Force

                Write-Step "Credentials updated. Retrying email test..."

                try {
                    $credential2 = New-Object System.Management.Automation.PSCredential(
                        $script:smtpUser,
                        $secPass2
                    )
                    Send-MailMessage `
                        -From $script:smtpUser `
                        -To $DefaultEmailTo `
                        -Subject "PC Masterclass - Deployment Test - $computerName" `
                        -Body "Email test successful for $computerName at $(Get-Date). This machine is ready to send maintenance reports." `
                        -SmtpServer "smtp.gmail.com" `
                        -Port 587 `
                        -UseSsl `
                        -Credential $credential2 `
                        -ErrorAction Stop

                    Write-OK "Test email sent successfully on retry!"
                    return $true
                } catch {
                    Write-Fail "Email still failing: $($_.Exception.Message)"
                    Write-Warn "Continuing with deployment - you'll need to fix email credentials later"
                    return $false
                }
            }
        }

        Write-Warn "Continuing with deployment - email may not work until credentials are fixed"
        return $false
    }
}

# ============================================================================
# SCHEDULED TASK SETUP
# ============================================================================
function Set-MaintenanceSchedule {
    Write-Section "SCHEDULED MAINTENANCE SETUP"
    Write-Host ""
    Write-Host "  The maintenance script can run automatically on a schedule." -ForegroundColor White
    Write-Host "  It will run unattended, generate a report, and email it to you." -ForegroundColor White
    Write-Host ""
    Write-Host "  Default: Every 90 days at $DefaultRunTime" -ForegroundColor White
    Write-Host "  Reports emailed to: $DefaultEmailTo" -ForegroundColor White
    Write-Host ""

    $setupSchedule = Read-Host "  Set up scheduled maintenance? (Y/N)"

    if ($setupSchedule -notmatch '^[Yy]') {
        Write-Warn "Skipped - maintenance will only run when triggered manually"
        return
    }

    # Ask for run time
    Write-Host ""
    $timeInput = Read-Host "  Run time (press Enter for $DefaultRunTime)"
    $runTime = if ($timeInput) { $timeInput } else { $DefaultRunTime }

    # Ask for frequency
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
        default { $DefaultFrequencyDays }
    }
    $freqLabel = switch ($frequencyDays) {
        30 { "monthly (every 30 days)" }
        60 { "bi-monthly (every 60 days)" }
        90 { "quarterly (every 90 days)" }
    }

    # Ask for start date (default to tomorrow so the first run fires at the scheduled time)
    Write-Host ""
    $tomorrowStr = (Get-Date).AddDays(1).ToString("dd/MM/yyyy")
    $dateInput = Read-Host "  Start date dd/MM/yyyy (press Enter for tomorrow, $tomorrowStr)"
    if ($dateInput) {
        try {
            $startDate = [datetime]::ParseExact($dateInput, "dd/MM/yyyy", $null)
        } catch {
            Write-Warn "Invalid date format - using tomorrow"
            $startDate = (Get-Date).AddDays(1).Date
        }
    } else {
        $startDate = (Get-Date).AddDays(1).Date
    }
    # Combine start date with chosen run time
    $startDateTime = $startDate.Add([datetime]::Parse($runTime).TimeOfDay)

    # Ask for email recipient
    Write-Host ""
    $emailInput = Read-Host "  Email reports to (press Enter for $DefaultEmailTo)"
    $emailTo = if ($emailInput) { $emailInput } else { $DefaultEmailTo }

    Write-Step "Creating scheduled task..."

    try {
        # Build the command that the scheduled task will run
        # -WindowStyle Hidden prevents a PowerShell window flashing on screen
        $scriptArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -EmailTo `"$emailTo`""

        # Create the scheduled task action
        $action = New-ScheduledTaskAction `
            -Execute "powershell.exe" `
            -Argument $scriptArgs

        # Create a daily trigger with the chosen day interval, starting on the chosen date
        $trigger = New-ScheduledTaskTrigger `
            -Daily `
            -DaysInterval $frequencyDays `
            -At $startDateTime

        # Task settings - try full settings first, fall back for older PowerShell versions
        # (Some older systems don't support -AllowStartIfOnBatteries)
        $settings = $null
        try {
            $settings = New-ScheduledTaskSettingsSet `
                -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries `
                -StartWhenAvailable `
                -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
                -MultipleInstances IgnoreNew
        } catch {
            Write-Warn "Advanced task settings not supported on this system - using defaults"
            try {
                $settings = New-ScheduledTaskSettingsSet `
                    -StartWhenAvailable `
                    -ExecutionTimeLimit (New-TimeSpan -Hours 4)
            } catch {
                Write-Warn "Task settings not available - registering without custom settings"
            }
        }

        # Register the task to run as SYSTEM (no password needed, runs elevated)
        $registerParams = @{
            TaskName    = $TaskName
            Action      = $action
            Trigger     = $trigger
            User        = "SYSTEM"
            RunLevel    = "Highest"
            Description = "PC Masterclass automated maintenance - runs $freqLabel and emails report to $emailTo"
            Force       = $true
        }
        if ($settings) {
            $registerParams.Settings = $settings
        }
        Register-ScheduledTask @registerParams | Out-Null

        # ---- VERIFY the DaysInterval actually stuck ----
        # Some Windows builds silently ignore -DaysInterval and default to 1 (daily).
        # If that happened, fix it by patching the task XML directly.
        $verifyTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        $verifyTrigger = $verifyTask.Triggers | Select-Object -First 1
        $actualInterval = 1
        if ($verifyTrigger -and $verifyTrigger.PSObject.Properties['DaysInterval']) {
            $actualInterval = $verifyTrigger.DaysInterval
        }

        if ($actualInterval -ne $frequencyDays) {
            Write-Warn "Windows set interval to $actualInterval day(s) instead of $frequencyDays — fixing via XML..."
            try {
                $xml = Export-ScheduledTask -TaskName $TaskName
                # The XML has <DaysInterval>1</DaysInterval> — replace with the correct value
                $xml = $xml -replace '<DaysInterval>\d+</DaysInterval>', "<DaysInterval>$frequencyDays</DaysInterval>"
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                Register-ScheduledTask -TaskName $TaskName -Xml $xml -User "SYSTEM" -Force | Out-Null
                Write-OK "Fixed: interval corrected to every $frequencyDays days via XML"
            } catch {
                Write-Fail "Could not fix interval via XML: $($_.Exception.Message)"
                Write-Warn "You may need to manually set the repeat interval in Task Scheduler"
            }
        } else {
            Write-OK "Verified: task interval correctly set to every $frequencyDays days"
        }

        Write-OK "Scheduled task created: $TaskName"
        Write-OK "Schedule: $freqLabel at $runTime"
        Write-OK "First scheduled run: $($startDateTime.ToString('dd MMM yyyy')) at $runTime"
        Write-OK "Reports will be emailed to: $emailTo"
        Write-OK "Task runs as SYSTEM with elevated privileges"

        # Trigger an immediate first run if start date is today
        if ($startDate.Date -le (Get-Date).Date) {
            Write-Step "Starting first maintenance run in background..."
            try {
                Start-ScheduledTask -TaskName $TaskName
                Write-OK "First run triggered - report will be emailed to $emailTo when complete (10-15 mins)"
            } catch {
                Write-Warn "Could not trigger first run automatically: $($_.Exception.Message)"
                Write-Warn "The task will run at its next scheduled time, or start it manually from Task Scheduler"
            }
        } else {
            Write-OK "First run will occur on $($startDate.ToString('dd MMM yyyy')) at $runTime"
        }

    } catch {
        Write-Fail "Failed to create scheduled task: $($_.Exception.Message)"
        Write-Warn "You can create it manually later via Task Scheduler"
    }
}

# ============================================================================
# DISPLAY SUMMARY
# ============================================================================
function Show-Summary {
    param([bool]$Success)

    Write-Host ""
    if ($Success) {
        $summary = @"

    ====================================================
       DEPLOYMENT COMPLETE
    ====================================================

       Maintenance script is ready at:
       $ScriptPath

       TO RUN MAINTENANCE:
       powershell -ExecutionPolicy Bypass -File "$ScriptPath"

       The script will auto-update from GitHub each time
       it runs, so this machine will always get the latest
       version.

    ====================================================
"@
        Write-Host $summary -ForegroundColor Green
    } else {
        $summary = @"

    ====================================================
       DEPLOYMENT INCOMPLETE - SEE ERRORS ABOVE
    ====================================================
       Some steps did not complete successfully.
       Please resolve the issues and run this script again.
    ====================================================
"@
        Write-Host $summary -ForegroundColor Red
    }
}

# ============================================================================
# MAIN EXECUTION
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

Show-Banner

$success = $true

# Step 1: Pre-flight checks
if (-not (Test-Prerequisites)) {
    Write-Fail "Pre-flight checks failed - resolve issues before continuing"
    Show-Summary -Success $false
    Read-Host "`n  Press Enter to close"
    return
}

# Step 2: Create folder structure
if (-not (New-FolderStructure)) {
    $success = $false
}

# Step 3: Download maintenance script
if ($success) {
    if (-not (Get-MaintenanceScript)) {
        $success = $false
    }
}

# Step 4: Verify installation
if ($success) {
    if (-not (Test-Installation)) {
        $success = $false
    }
}

# Step 5: Display machine info for Rollout Tracker
if ($success) {
    Get-MachineInfo
}

# Step 6: Email credential setup (optional, interactive)
if ($success) {
    Set-EmailCredentials
}

# Step 7: Test email sending (verifies credentials work before scheduling)
if ($success) {
    Test-EmailSend
}

# Step 8: Scheduled task setup (optional, interactive)
if ($success) {
    Set-MaintenanceSchedule
}

# Show final summary
Show-Summary -Success $success

Read-Host "`n  Press Enter to close"
