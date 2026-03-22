<#
.SYNOPSIS
    PC Masterclass - Email Credential Setup (Standalone)
.DESCRIPTION
    Sets up encrypted SMTP credentials on a client machine so the maintenance
    script can email reports. Creates both DPAPI (interactive user) and AES
    (SYSTEM scheduled task) credential files, then sends a test email.

    This is extracted from the full onboarding script for use during Tactical RMM
    rollouts where the agent is deployed first and credentials are set up separately.

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.0.0
    Date:    2026-03-17

    USAGE (paste into an elevated PowerShell prompt on the client machine):
      irm https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/Setup-EmailCreds-PCMasterclass.ps1 | iex

    OR run locally:
      powershell -ExecutionPolicy Bypass -File "C:\Teamviewer\Setup-EmailCreds-PCMasterclass.ps1"
#>

param(
    [string]$ClientName = ""
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$BaseDir   = "C:\Teamviewer"
$ConfigDir = Join-Path $BaseDir "Config"
$DefaultEmailTo = "reports@pcmasterclass.com.au"

# ============================================================================
# DISPLAY HELPERS
# ============================================================================
function Write-Step { param([string]$Message); Write-Host "  [*] $Message" -ForegroundColor White }
function Write-OK   { param([string]$Message); Write-Host "  [OK] $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message); Write-Host "  [!!] $Message" -ForegroundColor Yellow }
function Write-Fail { param([string]$Message); Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "  --- $Title ---" -ForegroundColor Cyan
}

# ============================================================================
# MAIN
# ============================================================================
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "    PC Masterclass - Email Credential Setup" -ForegroundColor Cyan
Write-Host "    v1.0.0" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan

# Ensure directories exist
Write-Section "CHECKING DIRECTORIES"
foreach ($dir in @($BaseDir, $ConfigDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-OK "Created: $dir"
    } else {
        Write-OK "Exists: $dir"
    }
}

# Check for existing credentials
$existingCred = Test-Path (Join-Path $ConfigDir "smtp-cred.xml")
$existingAes  = Test-Path (Join-Path $ConfigDir "smtp-config.xml")
if ($existingCred -or $existingAes) {
    Write-Warn "Existing credential files found - they will be overwritten"
}

# Gather credentials
Write-Section "EMAIL CREDENTIAL SETUP"
Write-Host ""
Write-Host "  Enter the SMTP credentials for sending maintenance reports." -ForegroundColor White
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
    Write-Fail "No password entered - exiting"
    Read-Host "`n  Press Enter to close"
    exit 1
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
    Write-OK "DPAPI credentials saved"

    # Save AES-encrypted fallback for SYSTEM scheduled task
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
        Write-OK "Credential file permissions restricted"
    } catch {
        Write-Warn "Could not restrict credential file permissions: $_"
    }

    Write-OK "AES fallback credentials saved"

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

} catch {
    Write-Fail "Failed to save credentials: $($_.Exception.Message)"
    Read-Host "`n  Press Enter to close"
    exit 1
}

# ============================================================================
# EMAIL SEND TEST
# ============================================================================
Write-Section "EMAIL SEND TEST"
Write-Step "Sending test email to $DefaultEmailTo..."

try {
    $secPass = ConvertTo-SecureString $plainPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential(
        $smtpUser,
        $secPass
    )

    $computerName = $env:COMPUTERNAME
    if ($ClientName) {
        $nameParts = $ClientName -split '\s+', 2
        $clientDisplay = if ($nameParts.Count -ge 2) { "$($nameParts[0].ToUpper()), $($nameParts[1])" } else { $ClientName }
    } else {
        $clientDisplay = $computerName
    }
    Send-MailMessage `
        -From $smtpUser `
        -To $DefaultEmailTo `
        -Subject "$clientDisplay - Email credential test of $computerName" `
        -Body "Email credential setup successful for $computerName at $(Get-Date). This machine is ready to send maintenance reports." `
        -SmtpServer "smtp.gmail.com" `
        -Port 587 `
        -UseSsl `
        -Credential $credential `
        -ErrorAction Stop

    Write-OK "Test email sent successfully to $DefaultEmailTo"
    Write-OK "Check your inbox to confirm delivery"

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
            $secPass2 = ConvertTo-SecureString $retryPlain -AsPlainText -Force

            # Re-save DPAPI credential
            @{
                SmtpUser   = $smtpUser
                SmtpServer = "smtp.gmail.com"
                SmtpPort   = 587
                EmailFrom  = "reports@pcmasterclass.com.au"
                Credential = New-Object System.Management.Automation.PSCredential(
                    $smtpUser,
                    $secPass2
                )
            } | Export-Clixml -Path (Join-Path $ConfigDir "smtp-cred.xml") -Force

            # Re-save AES credential
            $aesKey = New-Object byte[] 32
            [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($aesKey)
            Set-Content -Path (Join-Path $ConfigDir "smtp.key") -Value $aesKey -Encoding Byte -Force
            $enc2 = ConvertFrom-SecureString $secPass2 -Key $aesKey
            @{
                SmtpUser      = $smtpUser
                SmtpServer    = "smtp.gmail.com"
                SmtpPort      = 587
                EmailFrom     = "reports@pcmasterclass.com.au"
                EncryptedPass = $enc2
            } | Export-Clixml -Path (Join-Path $ConfigDir "smtp-config.xml") -Force

            Write-Step "Credentials updated. Retrying email test..."

            try {
                $credential2 = New-Object System.Management.Automation.PSCredential(
                    $smtpUser,
                    $secPass2
                )
                Send-MailMessage `
                    -From $smtpUser `
                    -To $DefaultEmailTo `
                    -Subject "$clientDisplay - Email credential test of $computerName" `
                    -Body "Email credential setup successful for $computerName at $(Get-Date). This machine is ready to send maintenance reports." `
                    -SmtpServer "smtp.gmail.com" `
                    -Port 587 `
                    -UseSsl `
                    -Credential $credential2 `
                    -ErrorAction Stop

                Write-OK "Test email sent successfully on retry!"
            } catch {
                Write-Fail "Email still failing: $($_.Exception.Message)"
                Write-Warn "Credentials have been saved - you may need to fix the App Password later"
            }
        }
    } else {
        Write-Warn "Credentials saved but email test was not successful"
        Write-Warn "Run Test-Email-PCMasterclass.ps1 later to diagnose"
    }
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Section "SETUP COMPLETE"
Write-Host ""
Write-Host "  Credential files in: $ConfigDir" -ForegroundColor White
Write-Host "    smtp-cred.xml    (DPAPI - interactive runs)" -ForegroundColor Gray
Write-Host "    smtp-config.xml  (AES - SYSTEM scheduled task)" -ForegroundColor Gray
Write-Host "    smtp.key         (AES decryption key)" -ForegroundColor Gray
Write-Host ""
Write-Host "  The maintenance script will use these automatically." -ForegroundColor White
Write-Host ""

Read-Host "  Press Enter to close"
