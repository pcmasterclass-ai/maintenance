#!/bin/bash
# ============================================================================
# PC Master Class — macOS Maintenance Script Installer
# Version: 1.1.0
# ============================================================================
# Installs a lightweight bundled Python runtime from python-build-standalone so
# client Macs do NOT need Apple Command Line Developer Tools / Xcode.
# ============================================================================

set -euo pipefail

SCRIPT_URL="https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/pcm_mac_maintenance.py"
INSTALL_DIR="$HOME/Library/PCMasterClass"
REPORT_DIR="$INSTALL_DIR/Reports"
LOG_DIR="$INSTALL_DIR/Logs"
RUNTIME_DIR="$INSTALL_DIR/python-runtime"
PYTHON_RUNTIME_DIR="$RUNTIME_DIR/python"
PYTHON_BIN="$PYTHON_RUNTIME_DIR/bin/python3"
LAUNCHAGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST_NAME="com.pcmasterclass.maintenance.agent.plist"
SCRIPT_NAME="pcm_mac_maintenance.py"
SCRIPT_PATH="$INSTALL_DIR/$SCRIPT_NAME"

# python-build-standalone, stripped install-only builds. These are ~27 MB
# downloads and avoid the 10-20+ GB Apple Developer Tools/Xcode prompt triggered
# by /usr/bin/python3 on Macs where Python is only an Apple stub.
PYTHON_VERSION="3.11.15"
PYTHON_BUILD_TAG="20260504"
PYTHON_ARM64_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20260504/cpython-3.11.15%2B20260504-aarch64-apple-darwin-install_only_stripped.tar.gz"
PYTHON_X86_64_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20260504/cpython-3.11.15%2B20260504-x86_64-apple-darwin-install_only_stripped.tar.gz"

printf '%s\n' "=========================================="
printf '%s\n' "PC Master Class macOS Maintenance"
printf '%s\n' "Installer v1.1.0"
printf '%s\n' "=========================================="
printf '\n'

# When this installer is run via curl ... | bash, stdin is the downloaded script
# stream rather than the keyboard. Read prompts from the controlling terminal.
TTY_INPUT="/dev/tty"
if [ ! -r "$TTY_INPUT" ]; then
    TTY_INPUT="/dev/stdin"
fi

need_cmd() {
    if ! command -v "$1" >/dev/null 2>&1; then
        echo "ERROR: Required command not found: $1"
        exit 1
    fi
}

need_cmd curl
need_cmd tar
need_cmd uname
need_cmd df

OS_MAJOR=$(sw_vers -productVersion | cut -d. -f1)
if [ "$OS_MAJOR" -lt 11 ]; then
    echo "WARNING: macOS 11 (Big Sur) or later is strongly recommended."
    read -r -p "Press Enter to continue or Ctrl-C to abort..." < "$TTY_INPUT"
fi

mkdir -p "$INSTALL_DIR" "$REPORT_DIR" "$LOG_DIR" "$LAUNCHAGENTS_DIR" "$RUNTIME_DIR"

install_python_runtime() {
    if [ -x "$PYTHON_BIN" ]; then
        if "$PYTHON_BIN" - <<'PY' >/dev/null 2>&1
import sys
raise SystemExit(0 if sys.version_info[:2] == (3, 11) else 1)
PY
        then
            echo "[+] Bundled Python already installed: $($PYTHON_BIN --version 2>&1)"
            return
        fi
        echo "[!] Existing bundled Python is not the expected version; reinstalling."
        rm -rf "$PYTHON_RUNTIME_DIR"
    fi

    ARCH="$(uname -m)"
    case "$ARCH" in
        arm64)
            RUNTIME_URL="$PYTHON_ARM64_URL"
            ;;
        x86_64)
            RUNTIME_URL="$PYTHON_X86_64_URL"
            ;;
        *)
            echo "ERROR: Unsupported Mac architecture: $ARCH"
            exit 1
            ;;
    esac

    # Keep a conservative floor; the runtime is about 27 MB compressed and
    # typically well under 200 MB installed, but extraction needs temporary room.
    AVAILABLE_KB=$(df -Pk "$HOME" | awk 'NR==2 {print $4}')
    REQUIRED_KB=512000
    if [ "${AVAILABLE_KB:-0}" -lt "$REQUIRED_KB" ]; then
        echo "ERROR: Not enough free disk space for lightweight Python runtime."
        echo "Available: $((AVAILABLE_KB / 1024)) MB; required: $((REQUIRED_KB / 1024)) MB."
        exit 1
    fi

    TMP_TGZ="$(mktemp /tmp/pcm-python-runtime.XXXXXX.tgz)"
    TMP_DIR="$(mktemp -d /tmp/pcm-python-runtime.XXXXXX)"
    cleanup_runtime_tmp() {
        rm -f "$TMP_TGZ"
        rm -rf "$TMP_DIR"
    }
    trap cleanup_runtime_tmp EXIT

    echo "[+] Downloading lightweight Python runtime ($PYTHON_VERSION, $ARCH, ~27 MB)..."
    curl -fL --retry 3 --connect-timeout 20 --max-time 600 "$RUNTIME_URL" -o "$TMP_TGZ"

    echo "[+] Extracting bundled Python runtime..."
    tar -xzf "$TMP_TGZ" -C "$TMP_DIR"
    rm -rf "$PYTHON_RUNTIME_DIR"
    mv "$TMP_DIR/python" "$PYTHON_RUNTIME_DIR"
    chmod -R u+rwX,go+rX "$PYTHON_RUNTIME_DIR"
    xattr -cr "$PYTHON_RUNTIME_DIR" 2>/dev/null || true

    if [ ! -x "$PYTHON_BIN" ]; then
        echo "ERROR: Bundled Python did not install correctly: $PYTHON_BIN missing."
        exit 1
    fi
    echo "[+] Bundled Python ready: $($PYTHON_BIN --version 2>&1)"
}

install_maintenance_script() {
    echo "[+] Downloading maintenance script..."
    curl -fsSL "$SCRIPT_URL?cachebust=$(date +%s)" -o "$SCRIPT_PATH"
    chmod +x "$SCRIPT_PATH"
    xattr -cr "$SCRIPT_PATH" 2>/dev/null || true
    if [ ! -s "$SCRIPT_PATH" ]; then
        echo "ERROR: Failed to download the maintenance script."
        exit 1
    fi
    "$PYTHON_BIN" -m py_compile "$SCRIPT_PATH"
    echo "[+] Script installed to: $SCRIPT_PATH"
}

install_python_runtime
install_maintenance_script

echo ""
echo "========================================"
echo "APPLE SECURITY SETUP REQUIRED"
echo "========================================"
echo "The maintenance script needs Full Disk Access for complete results."
echo "If you have already granted Terminal Full Disk Access and relaunched Terminal, press Enter."
read -r -p "Press Enter once Full Disk Access is ready (or to continue with limited results)..." < "$TTY_INPUT"

echo ""
echo "========================================"
echo "SMTP CREDENTIAL SETUP"
echo "========================================"
echo "Use the Gmail App Password from 1Password; do not use the regular account password."
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
# Gmail displays app passwords in grouped chunks; store/send without spaces.
SMTP_PASS="${SMTP_PASS//[[:space:]]/}"

read -r -p "From email [maintenance-reports@pcmasterclass.com.au]: " EMAIL_FROM < "$TTY_INPUT"
EMAIL_FROM=${EMAIL_FROM:-maintenance-reports@pcmasterclass.com.au}
read -r -p "Recipient email [reports@pcmasterclass.com.au]: " EMAIL_TO < "$TTY_INPUT"
EMAIL_TO=${EMAIL_TO:-reports@pcmasterclass.com.au}

CURRENT_USER=$(whoami)

echo "[+] Saving SMTP credentials to macOS Keychain..."
"$PYTHON_BIN" "$INSTALL_DIR/$SCRIPT_NAME" --save-credential \
    --smtp-user "$SMTP_USER" \
    --smtp-password "$SMTP_PASS" \
    --email-from "$EMAIL_FROM" \
    --email-to "$EMAIL_TO"
echo "[+] Credentials saved to macOS Keychain securely."

# Install fallback LaunchAgent. Tactical RMM is the preferred scheduler after
# onboarding, and this LaunchAgent may be removed after Tactical task verification.
echo ""
echo "[+] Installing fallback LaunchAgent for quarterly execution..."
cat > "$LAUNCHAGENTS_DIR/$PLIST_NAME" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.pcmasterclass.maintenance.agent</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON_BIN</string>
        <string>$INSTALL_DIR/$SCRIPT_NAME</string>
        <string>--email-to</string>
        <string>$EMAIL_TO</string>
        <string>--client-name</string>
        <string>$CURRENT_USER</string>
    </array>
    <key>StartCalendarInterval</key>
    <array>
        <dict><key>Month</key><integer>1</integer><key>Day</key><integer>15</integer><key>Hour</key><integer>1</integer><key>Minute</key><integer>0</integer></dict>
        <dict><key>Month</key><integer>4</integer><key>Day</key><integer>15</integer><key>Hour</key><integer>1</integer><key>Minute</key><integer>0</integer></dict>
        <dict><key>Month</key><integer>7</integer><key>Day</key><integer>15</integer><key>Hour</key><integer>1</integer><key>Minute</key><integer>0</integer></dict>
        <dict><key>Month</key><integer>10</integer><key>Day</key><integer>15</integer><key>Hour</key><integer>1</integer><key>Minute</key><integer>0</integer></dict>
    </array>
    <key>StandardOutPath</key>
    <string>$LOG_DIR/maintenance.out.log</string>
    <key>StandardErrorPath</key>
    <string>$LOG_DIR/maintenance.err.log</string>
</dict>
</plist>
PLIST

launchctl unload "$LAUNCHAGENTS_DIR/$PLIST_NAME" 2>/dev/null || true
launchctl load "$LAUNCHAGENTS_DIR/$PLIST_NAME" 2>/dev/null || true

echo "[+] LaunchAgent installed: $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "[+] Quarterly fallback schedule: Jan/Apr/Jul/Oct 15 at 1:00 AM"

echo ""
echo "========================================"
echo "INITIAL TEST RUN"
echo "========================================"
echo "Running a test scan now. This will generate a report and email it."
read -r -p "Press Enter to run the test scan..." < "$TTY_INPUT"

"$PYTHON_BIN" "$INSTALL_DIR/$SCRIPT_NAME" \
    --email-to "$EMAIL_TO" \
    --client-name "$CURRENT_USER" \
    --report-path "$REPORT_DIR"

echo ""
echo "========================================"
echo "INSTALLATION COMPLETE"
echo "========================================"
echo "Maintenance script: $SCRIPT_PATH"
echo "Bundled Python:     $PYTHON_BIN"
echo "Reports:            $REPORT_DIR"
echo "LaunchAgent:        $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "Logs:               $LOG_DIR"
echo ""
echo "To run manually:"
echo "  $PYTHON_BIN $SCRIPT_PATH --email-to $EMAIL_TO"
echo ""
echo "To uninstall:"
echo "  launchctl unload $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "  rm $LAUNCHAGENTS_DIR/$PLIST_NAME"
echo "  rm -rf $INSTALL_DIR"
