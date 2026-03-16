<#
.SYNOPSIS
    PC Masterclass - Email Diagnostic Test
.DESCRIPTION
    Quick diagnostic script to test whether a client machine can successfully
    send maintenance report emails. Checks credential files, decrypts them,
    and sends a test email to reports@pcmasterclass.com.au.

    Run this on any client machine where reports are not arriving.

.NOTES
    Author:  Paul - PC Masterclass
    Version: 1.0.0
    Date:    2026-03-17

    USAGE (paste into an elevated PowerShell prompt):
      irm https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/Test-Email-PCMasterclass.ps1 | iex

    Or if already downloaded:
      powershell -ExecutionPolicy Bypass -File "C:\Teamviewer\Test-Email-PCMasterclass.ps1"
#>

# ============================================================================
# CONFIGURATION
# ============================================================================
$BaseDir    = "C:\Teamviewer"
$ConfigDir  = Join-Path $BaseDir "Config"
$EmailTo    = "reports@pcmasterclass.com.au"
$SmtpServer = "smtp.gmail.com"
$SmtpPort   = 587

# ============================================================================
# HELPERS
# ============================================================================
function Write-OK   { param([string]$M) Write-Host "  [OK]   $M" -ForegroundColor Green }
function Write-FAIL { param([string]$M) Write-Host "  [FAIL] $M" -ForegroundColor Red }
function Write-WARN { param([string]$M) Write-Host "  [!!]   $M" -ForegroundColor Yellow }
function Write-INFO { param([string]$M) Write-Host "  [*]    $M" -ForegroundColor White }

# ============================================================================
# BANNER
# ============================================================================
Write-Host ""
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host "     PC Masterclass - Email Diagnostic Test v1.0.0" -ForegroundColor Cyan
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host ""

$computerName = $env:COMPUTERNAME
$isSystem = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM")
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

Write-INFO "Computer: $computerName"
Write-INFO "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-INFO "Admin: $isAdmin | SYSTEM: $isSystem"
Write-Host ""

# ============================================================================
# CHECK 1: Base directory exists
# ============================================================================
Write-Host "  --- Check 1: Directory Structure ---" -ForegroundColor Cyan
if (Test-Path $BaseDir) {
    Write-OK "Base directory exists: $BaseDir"
} else {
    Write-FAIL "Base directory missing: $BaseDir"
    Write-FAIL "The onboarding script has not been run on this machine."
    exit 1
}

if (Test-Path $ConfigDir) {
    Write-OK "Config directory exists: $ConfigDir"
} else {
    Write-FAIL "Config directory missing: $ConfigDir"
    Write-FAIL "Re-run the onboarding script to set up credentials."
    exit 1
}

# ============================================================================
# CHECK 2: Credential files exist
# ============================================================================
Write-Host ""
Write-Host "  --- Check 2: Credential Files ---" -ForegroundColor Cyan

$dpapiCred   = Join-Path $ConfigDir "smtp-cred.xml"
$aesKey      = Join-Path $ConfigDir "smtp.key"
$aesConfig   = Join-Path $ConfigDir "smtp-config.xml"

$hasDpapi  = Test-Path $dpapiCred
$hasAesKey = Test-Path $aesKey
$hasAesCfg = Test-Path $aesConfig

if ($hasDpapi)  { Write-OK "DPAPI credential file: $dpapiCred" }
else            { Write-WARN "DPAPI credential file missing (needed for interactive runs)" }

if ($hasAesKey) { Write-OK "AES key file: $aesKey" }
else            { Write-WARN "AES key file missing (needed for SYSTEM scheduled runs)" }

if ($hasAesCfg) { Write-OK "AES config file: $aesConfig" }
else            { Write-WARN "AES config file missing (needed for SYSTEM scheduled runs)" }

if (-not $hasDpapi -and -not $hasAesCfg) {
    Write-FAIL "No credential files found at all. Re-run the onboarding script."
    exit 1
}

# ============================================================================
# CHECK 3: Load and decrypt credentials
# ============================================================================
Write-Host ""
Write-Host "  --- Check 3: Decrypt Credentials ---" -ForegroundColor Cyan

$smtpUser     = $null
$smtpPassword = $null
$credSource   = $null

# Try DPAPI first (works for interactive runs under the user who created them)
if ($hasDpapi -and -not $isSystem) {
    Write-INFO "Attempting DPAPI decryption..."
    try {
        $cred = Import-Clixml -Path $dpapiCred
        $smtpUser = $cred.SmtpUser
        $smtpPassword = $cred.Credential.GetNetworkCredential().Password
        $credSource = "DPAPI"
        Write-OK "DPAPI decryption successful (user: $smtpUser)"
    } catch {
        Write-WARN "DPAPI decryption failed: $($_.Exception.Message)"
        Write-WARN "This is normal if you're running as a different user than who onboarded."
    }
}

# Fall back to AES (works for SYSTEM and any user)
if (-not $smtpPassword -and $hasAesKey -and $hasAesCfg) {
    Write-INFO "Attempting AES decryption..."
    try {
        $key = Get-Content -Path $aesKey -Encoding Byte -ErrorAction Stop
        $cfg = Import-Clixml -Path $aesConfig -ErrorAction Stop
        $smtpUser = $cfg.SmtpUser
        $secPass = $cfg.EncryptedPass | ConvertTo-SecureString -Key $key -ErrorAction Stop
        $smtpPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPass)
        )
        $credSource = "AES"
        Write-OK "AES decryption successful (user: $smtpUser)"
    } catch {
        Write-FAIL "AES decryption failed: $($_.Exception.Message)"
    }
}

if (-not $smtpPassword) {
    Write-FAIL "Could not decrypt credentials from any source."
    Write-FAIL "Re-run the onboarding script to set up fresh credentials."
    exit 1
}

Write-OK "Credential source: $credSource"
Write-OK "SMTP user: $smtpUser"
Write-OK "Password length: $($smtpPassword.Length) characters"

# Quick sanity check on password format
if ($smtpPassword.Contains(" ")) {
    Write-WARN "Password contains spaces! Gmail App Passwords should be 16 chars with NO spaces."
}
if ($smtpPassword.Length -ne 16) {
    Write-WARN "Password is $($smtpPassword.Length) chars. Gmail App Passwords are typically 16 chars."
}

# ============================================================================
# CHECK 4: Network connectivity to SMTP server
# ============================================================================
Write-Host ""
Write-Host "  --- Check 4: Network Connectivity ---" -ForegroundColor Cyan

Write-INFO "Testing connection to ${SmtpServer}:${SmtpPort}..."
try {
    $tcpTest = Test-NetConnection -ComputerName $SmtpServer -Port $SmtpPort -WarningAction SilentlyContinue
    if ($tcpTest.TcpTestSucceeded) {
        Write-OK "TCP connection to ${SmtpServer}:${SmtpPort} successful"
    } else {
        Write-FAIL "Cannot reach ${SmtpServer}:${SmtpPort}"
        Write-FAIL "Firewall or network issue is blocking SMTP traffic."
        Write-WARN "Check if the machine's firewall or router is blocking outbound port $SmtpPort."
        exit 1
    }
} catch {
    Write-WARN "Could not test network connection: $($_.Exception.Message)"
    Write-WARN "Proceeding with email test anyway..."
}

# ============================================================================
# CHECK 5: Send test email
# ============================================================================
Write-Host ""
Write-Host "  --- Check 5: Send Test Email ---" -ForegroundColor Cyan

Write-INFO "Sending test email to $EmailTo..."
Write-INFO "From: $smtpUser via $SmtpServer"

try {
    $secPass = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($smtpUser, $secPass)

    $timestamp = Get-Date -Format "dd MMM yyyy HH:mm:ss"
    Send-MailMessage `
        -From $smtpUser `
        -To $EmailTo `
        -Subject "Email Test - $computerName - $timestamp" `
        -Body "Email diagnostic test from $computerName at $timestamp.`n`nCredential source: $credSource`nSMTP user: $smtpUser`nRunning as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`n`nThis machine can send maintenance reports." `
        -SmtpServer $SmtpServer `
        -Port $SmtpPort `
        -UseSsl `
        -Credential $credential `
        -ErrorAction Stop

    Write-Host ""
    Write-OK "TEST EMAIL SENT SUCCESSFULLY"
    Write-OK "Check the Maintenance Reports label in Gmail for delivery."
    Write-Host ""

} catch {
    Write-Host ""
    Write-FAIL "EMAIL SEND FAILED"
    Write-FAIL "Error: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "  Common causes:" -ForegroundColor Yellow
    Write-Host "    - App Password was entered with spaces during onboarding" -ForegroundColor White
    Write-Host "    - App Password has been revoked in Google account settings" -ForegroundColor White
    Write-Host "    - Google account has 2FA disabled (App Passwords require 2FA)" -ForegroundColor White
    Write-Host "    - Firewall blocking outbound SMTP on port $SmtpPort" -ForegroundColor White
    Write-Host "    - Antivirus software intercepting SMTP connections" -ForegroundColor White
    Write-Host ""
    Write-WARN "To fix: re-run the onboarding script to set up fresh credentials."
    Write-Host ""
    exit 1
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host "     DIAGNOSTIC SUMMARY" -ForegroundColor Cyan
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host ""
Write-OK "Computer:       $computerName"
Write-OK "Credential:     $credSource ($smtpUser)"
Write-OK "SMTP:           ${SmtpServer}:${SmtpPort}"
Write-OK "Email to:       $EmailTo"
Write-OK "Result:         PASS - email sent"
Write-Host ""
Write-Host "  If the maintenance script is still not emailing reports," -ForegroundColor Yellow
Write-Host "  the issue may be in the script itself rather than email." -ForegroundColor Yellow
Write-Host "  Check C:\Teamviewer\Reports\ for local report files." -ForegroundColor Yellow
Write-Host ""
