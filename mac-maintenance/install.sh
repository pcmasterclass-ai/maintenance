#!/bin/bash
# ============================================================================
# PC Master Class — macOS Maintenance Script Installer
# Version: 1.0.2
# ============================================================================
# Usage:
#   chmod +x install.sh
#   ./install.sh
#
# This script:
#   1. Creates the directory structure in ~/Library/PCMasterClass/
#   2. Installs the main Python maintenance script
#   3. Installs a LaunchAgent for quarterly execution
#   4. Prompts for SMTP credential setup
#   5. Runs an initial test scan
# ============================================================================

set -e  # Exit on any error

SCRIPT_URL="https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/pcm_mac_maintenance.py"
UPDATE_URL="https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/install.sh"
LAUNCH_URL="https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/com.pcmasterclass.maintenance.agent.plist"

INSTALL_DIR="$HOME/Library/PCMasterClass"
REPORT_DIR="$HOME/Library/PCMasterClass/Reports"
LOG_DIR="$HOME/Library/PCMasterClass/Logs"
LAUNCHAGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST_NAME="com.pcmasterclass.maintenance.agent.plist"
SCRIPT_NAME="pcm_mac_maintenance.py"

echo "=========================================="
echo "PC Master Class macOS Maintenance"
echo "Installer v1.0.2"
echo "=========================================="
echo ""

# Check macOS version
OS_MAJOR=$(sw_vers -productVersion | cut -d. -f1)
if [ "$OS_MAJOR" -lt 11 ]; then
    echo "WARNING: macOS 11 (Big Sur) or later is strongly recommended."
    echo "Some features may not work correctly on macOS 10.x."
    echo "Press Enter to continue or Ctrl-C to abort."
    read
fi

# Create directories
echo "[+] Creating directories..."
mkdir -p "$INSTALL_DIR" "$REPORT_DIR" "$LOG_DIR" "$LAUNCHAGENTS_DIR"

# Download the main script
echo "[+] Downloading maintenance script..."
if command -v curl &> /dev/null; then
    curl -fsSL "$SCRIPT_URL" -o "$INSTALL_DIR/$SCRIPT_NAME"
elif command -v wget &> /dev/null; then
    wget -q "$SCRIPT_URL" -O "$INSTALL_DIR/$SCRIPT_NAME"
else
    echo "ERROR: curl or wget is required."
    exit 1
fi
chmod +x "$INSTALL_DIR/$SCRIPT_NAME"

# Verify download
if [ ! -s "$INSTALL_DIR/$SCRIPT_NAME" ]; then
    echo "ERROR: Failed to download the maintenance script."
    exit 1
fi

echo "[+] Script installed to: $INSTALL_DIR/$SCRIPT_NAME"

# ---------------------------------------------------------------------------
# Apple Security — Full Disk Access Check
# ---------------------------------------------------------------------------
echo ""
echo "========================================"
echo "APPLE SECURITY SETUP REQUIRED"
echo "========================================"
echo ""
echo "The maintenance script needs access to system information, user caches,"
echo "browser extensions, and network settings. macOS requires you to grant"
echo "Terminal (or the script runner) Full Disk Access."
echo ""
echo "Steps to grant access:"
echo "  1. Open System Settings (or System Preferences on older macOS)"
echo "  2. Go to Privacy & Security -> Full Disk Access"
echo "  3. Click the lock icon to make changes (authenticate)"
echo "  4. Click the '+' button and add 'Terminal' (or 'Python')"
echo "  5. Restart Terminal for the change to take effect"
echo ""
echo "If you skip this step, the script will still run but some checks"
echo "(e.g. browser extensions) will return limited results."
echo ""
# When this installer is run via `curl ... | bash`, stdin is the downloaded
# script rather than the user's keyboard. Read interactive answers from the
# controlling terminal so prompts still work correctly.
TTY_INPUT="/dev/tty"
if [ ! -r "$TTY_INPUT" ]; then
    TTY_INPUT="/dev/stdin"
fi
read -r -p "Press Enter once you've granted Full Disk Access (or to skip)..." < "$TTY_INPUT"

# ---------------------------------------------------------------------------
# Gatekeeper check
# ---------------------------------------------------------------------------
echo ""
echo "========================================"
echo "GATEKEEPER NOTE"
echo "========================================"
echo ""
echo "Since this script is downloaded from the internet, macOS may warn you"
echo "the first time you run it. If you see a Gatekeeper warning:"
echo ""
echo "  Option 1 (easiest): Right-click the script in Finder and select 'Open'"
echo "  Option 2: System Settings -> Privacy & Security -> Security -> Open Anyway"
echo ""
echo "The script itself is a plain text Python file — no code signing is needed."
echo "The warning only appears because it was downloaded (quarantine flag)."
echo ""

# ---------------------------------------------------------------------------
# SMTP Credentials Setup
# ---------------------------------------------------------------------------
echo ""
echo "========================================"
echo "SMTP CREDENTIAL SETUP"
echo "========================================"
echo ""
echo "The script sends maintenance reports via email."
echo "For Gmail, you MUST use an App Password (not your regular password)."
echo ""
echo "To create a Gmail App Password:"
echo "  1. Go to https://myaccount.google.com/apppasswords"
echo "  2. Sign in, select 'Mail', select 'Other (Custom name) = Mac Maintenance'"
echo "  3. Copy the 16-character password (e.g. abcd efgh ijkl mnop)"
echo "  4. Paste it below when prompted"
echo ""

read -r -p "SMTP login email [maintenance-reports@pcmasterclass.com.au]: " SMTP_USER < "$TTY_INPUT"
SMTP_USER=${SMTP_USER:-maintenance-reports@pcmasterclass.com.au}

SMTP_PASS=""
while [ -z "$SMTP_PASS" ]; do
    read -r -s -p "App Password: " SMTP_PASS < "$TTY_INPUT"
    echo ""
    if [ -z "$SMTP_PASS" ]; then
        echo "App Password cannot be blank. Please paste the Gmail App Password from 1Password."
    fi
done

read -r -p "From email [maintenance-reports@pcmasterclass.com.au]: " EMAIL_FROM < "$TTY_INPUT"
EMAIL_FROM=${EMAIL_FROM:-maintenance-reports@pcmasterclass.com.au}
read -r -p "Recipient email [reports@pcmasterclass.com.au]: " EMAIL_TO < "$TTY_INPUT"
EMAIL_TO=${EMAIL_TO:-reports@pcmasterclass.com.au}

if [ -z "$SMTP_USER" ] || [ -z "$SMTP_PASS" ]; then
    echo "WARNING: No credentials provided. You'll need to run credential setup manually:"
    echo "  python3 $INSTALL_DIR/$SCRIPT_NAME --save-credential --smtp-user YOUR_EMAIL --smtp-password YOUR_APP_PASSWORD"
else
    echo "[+] Saving SMTP credentials to macOS Keychain..."
    python3 "$INSTALL_DIR/$SCRIPT_NAME" --save-credential \
        --smtp-user "$SMTP_USER" \
        --smtp-password "$SMTP_PASS" \
        --email-from "$EMAIL_FROM" \
        --email-to "$EMAIL_TO"
    echo "[+] Credentials saved to macOS Keychain securely."
fi

# Install LaunchAgent
# ---------------------------------------------------------------------------
echo ""
echo "[+] Installing LaunchAgent for quarterly execution..."

# Get the user name for the plist
CURRENT_USER=$(whoami)

# Create the plist file with quarterly schedule (Jan 15, Apr 15, Jul 15, Oct 15)
cat > "$LAUNCHAGENTS_DIR/$PLIST_NAME" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.pcmasterclass.maintenance.agent</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>$INSTALL_DIR/$SCRIPT_NAME</string>
        <string>--email-to</string>
        <string>$EMAIL_TO</string>
        <string>--client-name</string>
        <string>$CURRENT_USER</string>
    </array>
    <key>StartCalendarInterval</key>
    <array>
        <!-- Q1: January 15 -->
        <dict>
            <key>Month</key>
            <integer>1</integer>
            <key>Day</key>
            <integer>15</integer>
            <key>Hour</key>
            <integer>1</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
        <!-- Q2: April 15 -->
        <dict>
            <key>Month</key>
            <integer>4</integer>
            <key>Day</key>
            <integer>15</integer>
            <key>Hour</key>
            <integer>1</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
        <!-- Q3: July 15 -->
        <dict>
            <key>Month</key>
            <integer>7</integer>
            <key>Day</key>
            <integer>15</integer>
            <key>Hour</key>
            <integer>1</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
        <!-- Q4: October 15 -->
        <dict>
            <key>Month</key>
            <integer>10</integer>
            <key>Day</key>
            <integer>15</integer>
            <key>Hour</key>
            <integer>1</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
    </array>
    <key>StandardOutPath</key>
    <string>$LOG_DIR/maintenance.out.log</string>
    <key>StandardErrorPath</key>
    <string>$LOG_DIR/maintenance.err.log</string>
</dict>
</plist>
PLIST

# Load the LaunchAgent
launchctl load "$LAUNCHAGENTS_DIR/$PLIST_NAME" 2>/dev/null || true

echo "[+] LaunchAgent installed: $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "[+] Quarterly schedule: Jan 15, Apr 15, Jul 15, Oct 15 at 1:00 AM"

# ---------------------------------------------------------------------------
# Placeholder for the separate LaunchAgent plist (kept in repo for reference)
# ---------------------------------------------------------------------------
# The file at mac-maintenance/com.pcmasterclass.maintenance.agent.plist
# is the canonical version.  This heredoc duplicates its behaviour for
# convenience during one-shot install.
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Initial test run
# ---------------------------------------------------------------------------
echo ""
echo "========================================"
echo "INITIAL TEST RUN"
echo "========================================"
echo ""
echo "Running a test scan now. This will generate a report and email it to you."
echo ""
read -r -p "Press Enter to run the test scan..." < "$TTY_INPUT"

TEST_REPORT="$REPORT_DIR/$(hostname)_Maintenance_$(date +%Y-%m-%d_%H-%M-%S).html"
python3 "$INSTALL_DIR/$SCRIPT_NAME" \
    --email-to "$EMAIL_TO" \
    --client-name "$CURRENT_USER" \
    --report-path "$REPORT_DIR"

echo ""
echo "========================================"
echo "INSTALLATION COMPLETE"
echo "========================================"
echo ""
echo "Maintenance script: $INSTALL_DIR/$SCRIPT_NAME"
echo "Reports:           $REPORT_DIR"
echo "LaunchAgent:       $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "Logs:              $LOG_DIR"
echo ""
echo "The script will run automatically every 3 months (Jan/Apr/Jul/Oct 15)."
echo "To run manually:"
echo "  python3 $INSTALL_DIR/$SCRIPT_NAME --email-to $EMAIL_TO"
echo ""
echo "To uninstall:"
echo "  launchctl unload $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "  rm $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "  rm -rf $INSTALL_DIR"
echo ""
