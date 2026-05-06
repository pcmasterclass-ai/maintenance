#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PC Master Class - macOS Maintenance Script
Version: 1.0.4
Author: Paul Benjamin
License: Proprietary

Runs a comprehensive maintenance checklist on client Apple Mac computers
and generates a branded HTML report with email delivery. Designed to be
installed via a simple shell script and run periodically via LaunchAgent.

USAGE:
  First-time credential setup (interactive):
    python3 pcm_mac_maintenance.py --save-credential \
        --smtp-user paul@pcmasterclass.com.au \
        --smtp-password YOUR_APP_PASSWORD \
        --email-from reports@pcmasterclass.com.au \
        --email-to reports@pcmasterclass.com.au

  Normal run (pulls credential from macOS Keychain):
    python3 pcm_mac_maintenance.py --email-to reports@pcmasterclass.com.au

  Options:
    --skip-updates      Skip Apple software update check
    --skip-smart        Skip disk health/SMART check (requires smartmontools)
    --client-name       Client name for report header
    --computer-name     Optional friendly computer/device name override
    --report-path       Custom path for report output
    --email-to          Send report via email
    --smtp-user         SMTP username
    --smtp-password     SMTP password (App Password for Gmail)
    --smtp-server       SMTP server (default: smtp.gmail.com)
    --smtp-port         SMTP port (default: 587)
    --email-from        Sender address
    --save-credential   Store credentials in macOS Keychain
    --skip-update-check Skip auto-update check from GitHub
"""

import argparse
import base64
import hashlib
import json
import os
import platform
import plistlib
import re
import shutil
import smtplib
import ssl
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# ============================================================================
# CONFIGURATION
# ============================================================================
SCRIPT_VERSION = "1.0.4"
UPDATE_URL = "https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/pcm_mac_maintenance.py"
UPDATE_API_URL = "https://api.github.com/repos/pcmasterclass-ai/maintenance/contents/mac-maintenance/pcm_mac_maintenance.py?ref=main"
UPDATE_TOKEN = ""  # Leave empty for public repos

# macOS Keychain service name for stored credentials
KEYCHAIN_SERVICE = "PCMasterClass Maintenance SMTP"

# ============================================================================
# LOGGING
# ============================================================================
class Logger:
    def __init__(self, log_file=None):
        self.log_file = log_file
        self.entries = []

    def log(self, message, level="INFO"):
        entry = f"[{level}] {datetime.now().strftime('%H:%M:%S')} - {message}"
        print(entry)
        self.entries.append(entry)
        if self.log_file:
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(entry + '\n')
            except Exception:
                pass

    def info(self, message):
        self.log(message, "INFO")

    def warn(self, message):
        self.log(message, "WARN")

    def error(self, message):
        self.log(message, "ERROR")


# ============================================================================
# KEYCHAIN CREDENTIAL MANAGEMENT
# ============================================================================
def keychain_store(account, password, smtp_server="smtp.gmail.com", smtp_port=587, email_from=""):
    """Store SMTP credentials in macOS Keychain."""
    if not email_from:
        email_from = account

    # Build the JSON payload
    payload = json.dumps({
        "smtp_user": account,
        "smtp_server": smtp_server,
        "smtp_port": str(smtp_port),
        "email_from": email_from,
        "version": SCRIPT_VERSION
    })

    # Store password using security command
    # We use the account name as the keychain 'account' field
    cmd = [
        "security", "add-generic-password",
        "-s", KEYCHAIN_SERVICE,
        "-a", account,
        "-w", password,
        "-T", "",  # Allow any application to access (simplifies script execution)
        "-U"       # Update if exists
    ]
    try:
        subprocess.run(cmd, capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError as e:
        # If item already exists, try delete then re-add
        if "already exists" in e.stderr.lower() or e.returncode == 45:
            subprocess.run(["security", "delete-generic-password", "-s", KEYCHAIN_SERVICE, "-a", account],
                             capture_output=True, text=True)
            subprocess.run(cmd, capture_output=True, text=True, check=True)
        else:
            raise RuntimeError(f"Failed to store credential in Keychain: {e.stderr}")

    # Store additional config (server, port, from) in a note
    note_cmd = [
        "security", "add-generic-password",
        "-s", KEYCHAIN_SERVICE + " Config",
        "-a", account,
        "-w", payload,
        "-T", "",
        "-U"
    ]
    try:
        subprocess.run(note_cmd, capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError:
        subprocess.run(["security", "delete-generic-password", "-s", KEYCHAIN_SERVICE + " Config", "-a", account],
                         capture_output=True, text=True)
        subprocess.run(note_cmd, capture_output=True, text=True, check=True)

    return True


def keychain_load(account):
    """Load SMTP credentials from macOS Keychain."""
    # Get password
    cmd = ["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", account, "-w"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        return None
    password = result.stdout.strip()

    # Get config
    cmd2 = ["security", "find-generic-password", "-s", KEYCHAIN_SERVICE + " Config", "-a", account, "-w"]
    result2 = subprocess.run(cmd2, capture_output=True, text=True)
    config = {}
    if result2.returncode == 0:
        try:
            config = json.loads(result2.stdout.strip())
        except json.JSONDecodeError:
            pass

    return {
        "smtp_user": account,
        "smtp_password": password,
        "smtp_server": config.get("smtp_server", "smtp.gmail.com"),
        "smtp_port": int(config.get("smtp_port", "587")),
        "email_from": config.get("email_from", account),
    }


def load_smtp_credentials(email_to=""):
    """Load SMTP credentials, preferring Paul's mailbox login for the reports@ alias."""
    accounts = []
    if email_to and email_to.lower() != "reports@pcmasterclass.com.au":
        accounts.append(email_to)
    # reports@ is an alias of Paul's mailbox. Prefer the real SMTP login first.
    accounts.extend(["paul@pcmasterclass.com.au", "reports@pcmasterclass.com.au"])

    seen = set()
    for account in accounts:
        key = account.lower()
        if key in seen:
            continue
        seen.add(key)
        loaded = keychain_load(account)
        if loaded:
            return loaded
    return None


def keychain_list_accounts():
    """List all accounts stored in Keychain for this service."""
    cmd = ["security", "dump-keychain"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    accounts = []
    if result.returncode == 0:
        # Parse output for our service
        for line in result.stdout.split('\n'):
            if KEYCHAIN_SERVICE in line and "svce" in line:
                # This is a heuristic; a more robust approach is needed
                pass
    # Simpler: try to find any password entries for our service
    cmd = ["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-g"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    # If we have an account-specific entry, the account name was passed
    # Since we don't know accounts a priori, we iterate common ones
    return []


# ============================================================================
# AUTO-UPDATE FROM GITHUB
# ============================================================================
def check_and_update():
    """Check for newer version on GitHub and self-update if found."""
    try:
        headers = {}
        if UPDATE_TOKEN:
            headers["Authorization"] = f"token {UPDATE_TOKEN}"

        context = ssl.create_default_context()
        remote_content = ""

        # Prefer GitHub's Contents API because raw.githubusercontent.com can be
        # stale behind CDN caches after a push. Fall back to the raw URL if the
        # API is unavailable.
        try:
            api_req = urllib.request.Request(UPDATE_API_URL, headers=headers)
            with urllib.request.urlopen(api_req, context=context, timeout=15) as response:
                payload = json.loads(response.read().decode('utf-8'))
            if payload.get("encoding") == "base64" and payload.get("content"):
                remote_content = base64.b64decode(payload["content"]).decode('utf-8')
        except Exception:
            remote_content = ""

        if not remote_content:
            req = urllib.request.Request(UPDATE_URL, headers=headers)
            with urllib.request.urlopen(req, context=context, timeout=15) as response:
                remote_content = response.read().decode('utf-8')

        # Extract remote version
        match = re.search(r'SCRIPT_VERSION\s*=\s*"([^"]+)"', remote_content)
        if not match:
            return False
        remote_version = match.group(1)

        try:
            from packaging import version as pv
            needs_update = pv.parse(remote_version) > pv.parse(SCRIPT_VERSION)
        except ImportError:
            needs_update = remote_version != SCRIPT_VERSION

        if not needs_update:
            # Even if versions match, check hash for hotfixes
            current_path = Path(sys.argv[0]).resolve()
            local_hash = hashlib.sha256(current_path.read_bytes()).hexdigest()
            remote_hash = hashlib.sha256(remote_content.encode('utf-8')).hexdigest()
            if local_hash != remote_hash:
                print(f"[UPDATE] Same version ({SCRIPT_VERSION}) but content differs — applying hotfix...")
                needs_update = True

        if needs_update:
            print(f"[UPDATE] New version available: {remote_version} (current: {SCRIPT_VERSION})")
            current_path = Path(sys.argv[0]).resolve()
            backup = current_path.with_suffix('.py.bak')
            shutil.copy2(current_path, backup)
            current_path.write_text(remote_content, encoding='utf-8')
            print(f"[UPDATE] Updated to v{remote_version}. Restarting...")
            os.execv(sys.executable, [sys.executable, str(current_path)] + sys.argv[1:])
            return True

        return False

    except Exception as e:
        print(f"[WARNING] Auto-update check failed: {e}")
        return False


# ============================================================================
# MAC HARDWARE DISPLAY NAME HELPERS
# ============================================================================
def _normalize_chip_name(value):
    """Return a readable Apple/Intel processor label from system_profiler/sysctl values."""
    if not value or value == "Unknown":
        return "Unknown"
    value = re.sub(r"\s+", " ", str(value)).strip()
    if value.startswith("Apple "):
        return value
    # system_profiler normally reports Apple Silicon as "Apple Mx...". For Intel,
    # keep the CPU brand but trim noisy clock-speed suffixes where practical.
    return value.replace("(R)", "").replace("(TM)", "").strip()


# High-value Apple Silicon model identifier mappings for report display names.
# Fallbacks still work if an unknown future identifier appears; the report will show
# the system_profiler model name plus processor rather than a guessed size/year.
MODEL_IDENTIFIER_DETAILS = {
    # MacBook Air
    "Mac14,2": {"friendly_model": "13-inch MacBook Air", "release_year": "2022"},
    "Mac14,15": {"friendly_model": "15-inch MacBook Air", "release_year": "2023"},
    "Mac15,12": {"friendly_model": "13-inch MacBook Air", "release_year": "2024"},
    "Mac15,13": {"friendly_model": "15-inch MacBook Air", "release_year": "2024"},
    "Mac16,12": {"friendly_model": "13-inch MacBook Air", "release_year": "2025"},
    "Mac16,13": {"friendly_model": "15-inch MacBook Air", "release_year": "2025"},

    # MacBook Pro — Apple Silicon generations where model identifier differentiates size in practice.
    "MacBookPro17,1": {"friendly_model": "13-inch MacBook Pro", "release_year": "2020"},
    "MacBookPro18,3": {"friendly_model": "14-inch MacBook Pro", "release_year": "2021"},
    "MacBookPro18,4": {"friendly_model": "14-inch MacBook Pro", "release_year": "2021"},
    "MacBookPro18,1": {"friendly_model": "16-inch MacBook Pro", "release_year": "2021"},
    "MacBookPro18,2": {"friendly_model": "16-inch MacBook Pro", "release_year": "2021"},
    "Mac14,5": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac14,9": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac14,6": {"friendly_model": "16-inch MacBook Pro", "release_year": "2023"},
    "Mac14,10": {"friendly_model": "16-inch MacBook Pro", "release_year": "2023"},
    "Mac15,3": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac15,6": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac15,8": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac15,10": {"friendly_model": "14-inch MacBook Pro", "release_year": "2023"},
    "Mac15,7": {"friendly_model": "16-inch MacBook Pro", "release_year": "2023"},
    "Mac15,9": {"friendly_model": "16-inch MacBook Pro", "release_year": "2023"},
    "Mac15,11": {"friendly_model": "16-inch MacBook Pro", "release_year": "2023"},
    "Mac16,1": {"friendly_model": "14-inch MacBook Pro", "release_year": "2024"},
    "Mac16,6": {"friendly_model": "14-inch MacBook Pro", "release_year": "2024"},
    "Mac16,8": {"friendly_model": "16-inch MacBook Pro", "release_year": "2024"},
    "Mac16,10": {"friendly_model": "16-inch MacBook Pro", "release_year": "2024"},

    # iMac
    "iMac21,1": {"friendly_model": "24-inch iMac", "release_year": "2021"},
    "iMac21,2": {"friendly_model": "24-inch iMac", "release_year": "2021"},
    "Mac15,4": {"friendly_model": "24-inch iMac", "release_year": "2023"},
    "Mac15,5": {"friendly_model": "24-inch iMac", "release_year": "2023"},
    "Mac16,2": {"friendly_model": "24-inch iMac", "release_year": "2024"},
    "Mac16,3": {"friendly_model": "24-inch iMac", "release_year": "2024"},

    # Mac mini / Studio / Pro
    "Macmini9,1": {"friendly_model": "Mac mini", "release_year": "2020"},
    "Mac14,3": {"friendly_model": "Mac mini", "release_year": "2023"},
    "Mac14,12": {"friendly_model": "Mac mini", "release_year": "2023"},
    "Mac16,11": {"friendly_model": "Mac mini", "release_year": "2024"},
    "Mac16,15": {"friendly_model": "Mac mini", "release_year": "2024"},
    "Mac13,1": {"friendly_model": "Mac Studio", "release_year": "2022"},
    "Mac13,2": {"friendly_model": "Mac Studio", "release_year": "2022"},
    "Mac14,13": {"friendly_model": "Mac Studio", "release_year": "2023"},
    "Mac14,14": {"friendly_model": "Mac Studio", "release_year": "2023"},
    "Mac16,9": {"friendly_model": "Mac Studio", "release_year": "2025"},
    "Mac14,8": {"friendly_model": "Mac Pro", "release_year": "2023"},
}


def build_hardware_display_name(system_info, computer_name_override=""):
    """Build a report-friendly computer label from Apple hardware details."""
    if computer_name_override:
        return computer_name_override

    model_id = system_info.get("Model") or system_info.get("ModelName") or ""
    details = MODEL_IDENTIFIER_DETAILS.get(model_id, {})
    friendly_model = details.get("friendly_model") or system_info.get("MachineName") or system_info.get("ModelName") or "Mac"
    processor = _normalize_chip_name(system_info.get("ProcessorName") or system_info.get("Chip") or system_info.get("CPU") or "")
    release_year = details.get("release_year", "")

    extra = ", ".join(part for part in (processor if processor != "Unknown" else "", release_year) if part)
    return f"{friendly_model} ({extra})" if extra else friendly_model


# ============================================================================
# SYSTEM INFO MODULE
# ============================================================================
def run_system_info():
    """Gather basic system information."""
    logger.info("Gathering system information...")
    result = {
        "Status": "PASS",
        "Hostname": platform.node(),
        "OS": platform.mac_ver()[0],
        "Build": platform.mac_ver()[2],
    }

    # Model
    try:
        model = subprocess.run(["sysctl", "-n", "hw.model"], capture_output=True, text=True, check=True).stdout.strip()
        result["Model"] = model
    except Exception:
        result["Model"] = "Unknown"

    # CPU
    try:
        cpu = subprocess.run(["sysctl", "-n", "machdep.cpu.brand_string"], capture_output=True, text=True, check=True).stdout.strip()
        cores = subprocess.run(["sysctl", "-n", "hw.physicalcpu"], capture_output=True, text=True, check=True).stdout.strip()
        threads = subprocess.run(["sysctl", "-n", "hw.logicalcpu"], capture_output=True, text=True, check=True).stdout.strip()
        result["CPU"] = f"{cpu} ({cores} cores / {threads} threads)"
    except Exception:
        result["CPU"] = "Unknown"

    if result.get("ProcessorName") in (None, "", "Unknown") and result.get("CPU"):
        result["ProcessorName"] = _normalize_chip_name(result.get("CPU", "Unknown").split("(")[0].strip())

    # RAM
    try:
        mem_bytes = int(subprocess.run(["sysctl", "-n", "hw.memsize"], capture_output=True, text=True, check=True).stdout.strip())
        result["RAM_GB"] = round(mem_bytes / (1024 ** 3), 1)
    except Exception:
        result["RAM_GB"] = "Unknown"

    # Serial
    try:
        serial = subprocess.run(["system_profiler", "SPHardwareDataType", "-json"], capture_output=True, text=True, check=True).stdout.strip()
        data = json.loads(serial)
        sp = data.get("SPHardwareDataType", [{}])[0]
        result["SerialNumber"] = sp.get("serial_number", "Unknown")
        result["Manufacturer"] = "Apple"
        # system_profiler uses different keys across macOS/CPU generations.
        result["MachineName"] = sp.get("machine_name", "")
        result["ModelName"] = sp.get("machine_name", sp.get("machine_model", "Unknown"))
        result["ModelIdentifier"] = sp.get("machine_model", result.get("Model", "Unknown"))
        result["ProcessorName"] = _normalize_chip_name(sp.get("chip_type") or sp.get("current_processor_speed") or sp.get("cpu_type", "Unknown"))
    except Exception:
        result["SerialNumber"] = "Unknown"
        result["Manufacturer"] = "Apple"
        result["MachineName"] = "Unknown"
        result["ModelName"] = "Unknown"
        result["ModelIdentifier"] = result.get("Model", "Unknown")
        result["ProcessorName"] = "Unknown"

    # Boot time / uptime
    try:
        boot_time = int(subprocess.run(["sysctl", "-n", "kern.boottime"], capture_output=True, text=True, check=True).stdout.strip().split("{")[1].split("}")[0].split("sec = ")[1].split(",")[0].strip())
        uptime_seconds = int(time.time()) - boot_time
        uptime_days = round(uptime_seconds / 86400, 1)
        result["UptimeDays"] = uptime_days
        result["LastBoot"] = datetime.fromtimestamp(boot_time).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        result["UptimeDays"] = "Unknown"
        result["LastBoot"] = "Unknown"

    # Battery health (if applicable)
    try:
        bat = subprocess.run(["pmset", "-g", "batt"], capture_output=True, text=True)
        if bat.returncode == 0 and "Battery" in bat.stdout:
            lines = bat.stdout.strip().split('\n')
            for line in lines:
                if "%;" in line:
                    pct = line.split("%;")[0].split()[-1] + "%"
                    result["Battery"] = pct
                    break

        battery_info = subprocess.run(["system_profiler", "SPPowerDataType", "-json"], capture_output=True, text=True, timeout=30)
        if battery_info.returncode == 0 and battery_info.stdout.strip():
            power = json.loads(battery_info.stdout)
            batteries = power.get("SPPowerDataType", [{}])[0].get("sppower_battery_health_info", [])
            if batteries:
                binfo = batteries[0]
                result["BatteryCycleCount"] = binfo.get("sppower_battery_cycle_count", "Unknown")
                result["BatteryCondition"] = binfo.get("sppower_battery_condition", "Unknown")
    except Exception:
        pass

    # FileVault status
    try:
        fv = subprocess.run(["fdesetup", "status"], capture_output=True, text=True, check=True)
        fv_out = fv.stdout.strip()
        if "FileVault is On" in fv_out:
            result["FileVault"] = "On"
        elif "FileVault is Off" in fv_out:
            result["FileVault"] = "Off"
        else:
            result["FileVault"] = fv_out
    except Exception:
        result["FileVault"] = "Unknown"

    # Network adapters
    result["NetworkAdapters"] = []
    try:
        iface = subprocess.run(["ifconfig"], capture_output=True, text=True, check=True).stdout
        current = {}
        for line in iface.split('\n'):
            if line.startswith('\\t') or line.startswith('    '):
                if 'inet ' in line and 'netmask' in line:
                    ip = line.split('inet ')[1].split()[0]
                    current['IPv4'] = ip
                if 'ether ' in line:
                    current['MAC'] = line.split('ether ')[1].split()[0]
            elif ':' in line and not line.startswith(' '):
                if current and current.get('IPv4'):
                    result["NetworkAdapters"].append(current)
                current = {"Name": line.split(':')[0], "IPv4": "", "MAC": "", "Gateway": "", "DNS": ""}
        if current and current.get('IPv4'):
            result["NetworkAdapters"].append(current)
    except Exception:
        pass

    return result


# ============================================================================
# DISK HEALTH MODULE
# ============================================================================
def run_disk_health():
    """Check disk health, usage, and SMART if smartmontools is installed."""
    logger.info("Checking disk health...")
    result = {"Status": "PASS", "Disks": [], "Volumes": []}

    # Disk info via diskutil
    try:
        disk_list = json.loads(subprocess.run(["diskutil", "list", "-json"], capture_output=True, text=True, check=True).stdout)
        for disk in disk_list.get("AllDisksAndPartitions", []):
            disk_id = disk.get("DeviceIdentifier", "unknown")
            size = disk.get("Size", 0)
            size_gb = round(size / (1024 ** 3), 1) if size else "?"
            try:
                info = json.loads(subprocess.run(["diskutil", "info", "-json", disk_id], capture_output=True, text=True, check=True).stdout)
                di = info.get("Disk", {})
                result["Disks"].append({
                    "Name": disk_id,
                    "Type": di.get("SolidState", False) and "SSD" or "HDD",
                    "SizeGB": size_gb,
                    "Protocol": di.get("Protocol", "Unknown"),
                    "SMART": "N/A"
                })
            except Exception:
                result["Disks"].append({"Name": disk_id, "Type": "?", "SizeGB": size_gb, "Protocol": "?", "SMART": "N/A"})
    except Exception as e:
        result["Status"] = "WARNING"
        result["Error"] = str(e)

    # Volumes
    try:
        df = subprocess.run(["df", "-h"], capture_output=True, text=True, check=True).stdout
        for line in df.split('\n')[1:]:
            parts = line.split()
            if len(parts) >= 9:
                filesystem = parts[0]
                # Skip virtual filesystems
                if filesystem in ("devfs", "map", "com.apple") or filesystem.startswith("com.apple"):
                    continue
                size = parts[1]
                used = parts[2]
                avail = parts[3]
                pct_str = parts[4]
                mount = parts[8]
                try:
                    pct = int(pct_str.replace('%', ''))
                except ValueError:
                    pct = 0
                warning = pct > 85
                result["Volumes"].append({
                    "Filesystem": filesystem,
                    "Mount": mount,
                    "Size": size,
                    "Used": used,
                    "Avail": avail,
                    "UsedPercent": pct,
                    "Warning": warning
                })
                if warning:
                    result["Status"] = "WARNING"
    except Exception as e:
        logger.error(f"Disk volume check failed: {e}")

    # SMART via smartmontools (optional)
    smart_path = shutil.which("smartctl")
    if smart_path:
        try:
            for disk in result["Disks"]:
                disk_id = disk["Name"]
                smart = subprocess.run([smart_path, "-a", f"/dev/{disk_id}"], capture_output=True, text=True)
                if smart.returncode in (0, 2):
                    # Parse SMART for overall health
                    if "PASSED" in smart.stdout:
                        disk["SMART"] = "PASSED"
                    elif "FAILED" in smart.stdout:
                        disk["SMART"] = "FAILED"
                        result["Status"] = "WARNING"
                    else:
                        disk["SMART"] = "Unknown"
        except Exception:
            pass
    else:
        result["SMARTNote"] = "smartmontools not installed. Install via 'brew install smartmontools' for detailed SMART data."

    return result


# ============================================================================
# SOFTWARE UPDATES MODULE
# ============================================================================
def run_software_updates():
    """Check for pending macOS and optionally Homebrew updates."""
    logger.info("Checking for software updates...")
    result = {"Status": "PASS", "PendingCount": 0, "Updates": [], "RebootRequired": False}

    try:
        updates = subprocess.run(["softwareupdate", "-l"], capture_output=True, text=True)
        out = updates.stdout + updates.stderr

        # Parse recommended updates
        recommended = []
        in_section = False
        for line in out.split('\n'):
            if "Recommended\\s*" in line or "Recommended" in line:
                in_section = True
                continue
            if in_section and line.strip().startswith("*"):
                recommended.append(line.strip().lstrip("* ").strip())
            elif in_section and not line.strip():
                break

        if recommended:
            for upd in recommended:
                result["Updates"].append({"Title": upd, "IsImportant": True})
            result["PendingCount"] = len(recommended)
            result["Status"] = "WARNING"
        else:
            # Check for no updates available
            if "No new software available" in out or "No updates available" in out:
                result["PendingCount"] = 0
            else:
                result["PendingCount"] = 0  # Assume none if we can't parse

        # Reboot required?
        if "restart" in out.lower() or "reboot" in out.lower():
            result["RebootRequired"] = True

    except Exception as e:
        result["Status"] = "ERROR"
        result["Error"] = str(e)

    # Homebrew check (optional)
    if shutil.which("brew"):
        try:
            brew = subprocess.run(["brew", "outdated"], capture_output=True, text=True)
            if brew.stdout.strip():
                brew_packages = brew.stdout.strip().split('\n')
                result["HomebrewUpdates"] = len(brew_packages)
                result["Status"] = "WARNING" if result["Status"] != "ERROR" else result["Status"]
        except Exception:
            pass

    return result


# ============================================================================
# BACKUP MODULE (iDrive + Time Machine)
# ============================================================================
def run_backup_check():
    """Check iDrive and/or Time Machine backup status."""
    logger.info("Checking backup status...")
    result = {
        "Status": "PASS",
        "iDrive": {"Installed": False},
        "TimeMachine": {"Configured": False}
    }

    # iDrive check
    idrive_app = Path("/Applications/iDrive.app")
    if idrive_app.exists():
        result["iDrive"]["Installed"] = True
        # Try to find iDrive service or status
        try:
            ps = subprocess.run(["pgrep", "-f", "iDrive"], capture_output=True, text=True)
            result["iDrive"]["Running"] = ps.returncode == 0 and ps.stdout.strip() != ""
        except Exception:
            result["iDrive"]["Running"] = False

        # Try to read iDrive logs or status
        # iDrive on macOS stores logs in ~/Library/Application Support/iDrive
        user_home = Path.home()
        idrive_logs = user_home / "Library/Application Support/iDrive"
        if idrive_logs.exists():
            result["iDrive"]["LogPath"] = str(idrive_logs)
            # Look for the most recent backup timestamp in log files
            latest = 0
            for log_file in idrive_logs.glob("*.log"):
                try:
                    mtime = log_file.stat().st_mtime
                    if mtime > latest:
                        latest = mtime
                except Exception:
                    pass
            if latest:
                result["iDrive"]["LastLog"] = datetime.fromtimestamp(latest).strftime("%Y-%m-%d %H:%M:%S")
    else:
        result["iDrive"]["Installed"] = False

    # Time Machine check
    try:
        # Destination info is the best signal that Time Machine is configured, even when no backup is mounted.
        dest = subprocess.run(["tmutil", "destinationinfo"], capture_output=True, text=True)
        if dest.returncode == 0 and dest.stdout.strip():
            result["TimeMachine"]["Configured"] = True
            result["TimeMachine"]["Destinations"] = dest.stdout.strip()

        tm_status = subprocess.run(["tmutil", "status"], capture_output=True, text=True)
        if tm_status.returncode == 0 and tm_status.stdout.strip():
            result["TimeMachine"]["StatusText"] = tm_status.stdout.strip()
            result["TimeMachine"]["Running"] = "Running = 1" in tm_status.stdout

        tm = subprocess.run(["tmutil", "latestbackup"], capture_output=True, text=True)
        if tm.returncode == 0 and tm.stdout.strip():
            result["TimeMachine"]["Configured"] = True
            path = tm.stdout.strip()
            result["TimeMachine"]["LastBackupPath"] = path
            # Extract date from path (e.g., /Volumes/.../2026-05-01-123456/Macintosh HD)
            match = re.search(r'(\\d{4}-\\d{2}-\\d{2}-\\d{2,6})', path)
            if match:
                dt_str = match.group(1)
                for fmt in ("%Y-%m-%d-%H%M%S", "%Y-%m-%d-%H"):
                    try:
                        result["TimeMachine"]["LastBackupDate"] = datetime.strptime(dt_str, fmt).strftime("%Y-%m-%d %H:%M:%S")
                        break
                    except ValueError:
                        pass
                else:
                    result["TimeMachine"]["LastBackupDate"] = "Unknown"
            else:
                result["TimeMachine"]["LastBackupDate"] = "Unknown"

        snapshots = subprocess.run(["tmutil", "listlocalsnapshots", "/"], capture_output=True, text=True)
        if snapshots.returncode == 0:
            snaps = [line.strip() for line in snapshots.stdout.splitlines() if line.strip()]
            result["TimeMachine"]["LocalSnapshotCount"] = len(snaps)
    except Exception as e:
        result["TimeMachine"]["Configured"] = False
        result["TimeMachine"]["Error"] = str(e)

    # Determine overall status
    if not result["iDrive"].get("Installed") and not result["TimeMachine"].get("Configured"):
        result["Status"] = "WARNING"
    elif result["TimeMachine"].get("Configured"):
        # Check if backup is old (> 7 days)
        try:
            last = result["TimeMachine"].get("LastBackupDate", "")
            if last and last != "Unknown":
                last_dt = datetime.strptime(last, "%Y-%m-%d %H:%M:%S")
                days_since = (datetime.now() - last_dt).days
                if days_since > 7:
                    result["Status"] = "WARNING"
                    result["TimeMachine"]["DaysSince"] = days_since
        except Exception:
            pass

    return result


# ============================================================================
# MALWAREBYTES MODULE
# ============================================================================
def run_malwarebytes_check():
    """Check Malwarebytes for Mac status."""
    logger.info("Checking Malwarebytes...")
    result = {"Status": "PASS", "Installed": False}

    # Check for Malwarebytes app
    mb_app = Path("/Applications/Malwarebytes.app")
    if not mb_app.exists():
        mb_app = Path("/Applications/Malwarebytes Endpoint Protection.app")

    if mb_app.exists():
        result["Installed"] = True
        result["ProductType"] = mb_app.name

        # Check if service is running
        try:
            ps = subprocess.run(["pgrep", "-f", "Malwarebytes"], capture_output=True, text=True)
            result["ServiceRunning"] = ps.returncode == 0 and ps.stdout.strip() != ""
        except Exception:
            result["ServiceRunning"] = False

        if not result["ServiceRunning"]:
            result["Status"] = "WARNING"
    else:
        result["Status"] = "INFO"
        result["Note"] = "Malwarebytes not installed on this Mac."

    return result


# ============================================================================
# XPROTECT / GATEKEEPER MODULE (replaces Windows Defender)
# ============================================================================
def run_security_check():
    """Check macOS security: Gatekeeper, XProtect, SIP."""
    logger.info("Checking macOS security status...")
    result = {"Status": "PASS", "Gatekeeper": "Unknown", "XProtect": "Unknown", "SIP": "Unknown"}

    # Gatekeeper
    try:
        gk = subprocess.run(["spctl", "--status"], capture_output=True, text=True)
        if "assessments enabled" in gk.stdout:
            result["Gatekeeper"] = "Enabled"
        elif "disabled" in gk.stdout:
            result["Gatekeeper"] = "Disabled"
            result["Status"] = "WARNING"
        else:
            result["Gatekeeper"] = gk.stdout.strip()
    except Exception:
        pass

    # XProtect version
    try:
        xp = subprocess.run(["system_profiler", "SPInstallHistoryDataType", "-json"], capture_output=True, text=True)
        if xp.returncode == 0:
            data = json.loads(xp.stdout)
            for item in data.get("SPInstallHistoryDataType", []):
                if "XProtect" in item.get("package_id", ""):
                    result["XProtect"] = item.get("package_version", "Unknown")
                    result["XProtectInstallDate"] = item.get("install_date", "Unknown")
                    break
    except Exception:
        pass

    # SIP (System Integrity Protection)
    try:
        sip = subprocess.run(["csrutil", "status"], capture_output=True, text=True)
        if "enabled" in sip.stdout.lower():
            result["SIP"] = "Enabled"
        elif "disabled" in sip.stdout.lower():
            result["SIP"] = "Disabled"
            result["Status"] = "WARNING"
        else:
            result["SIP"] = sip.stdout.strip()
    except Exception:
        pass

    return result


# ============================================================================
# LOG ERRORS MODULE
# ============================================================================
def run_log_errors(hours=24):
    """Check for actionable crash reports and high-signal system errors."""
    logger.info("Checking system logs for actionable errors...")
    result = {"Status": "PASS", "HoursChecked": hours, "TotalEvents": 0, "Events": [], "CrashReports": [], "KernelPanics": []}

    try:
        cutoff = time.time() - (hours * 3600)
        diagnostic_dirs = [
            Path("~/Library/Logs/DiagnosticReports").expanduser(),
            Path("/Library/Logs/DiagnosticReports"),
        ]

        for crash_reports_dir in diagnostic_dirs:
            if crash_reports_dir.exists():
                for f in crash_reports_dir.iterdir():
                    try:
                        if not f.is_file() or f.stat().st_mtime <= cutoff:
                            continue
                        name = f.name
                        lower = name.lower()
                        if "panic" in lower or name.endswith(".panic"):
                            result["KernelPanics"].append(name)
                        elif name.endswith(".crash"):
                            result["CrashReports"].append(name)
                    except Exception:
                        pass

        # Unified logs are very noisy on macOS. Only search for a few high-signal disk/APFS/kernel messages.
        since = (datetime.now() - __import__('datetime').timedelta(hours=hours)).strftime("%Y-%m-%d %H:%M:%S")
        predicate = '(eventType == logEvent) && (messageType == error) && (eventMessage CONTAINS[c] "I/O error" || eventMessage CONTAINS[c] "APFS" || eventMessage CONTAINS[c] "disk" || eventMessage CONTAINS[c] "panic")'
        cmd = ["log", "show", "--predicate", predicate, "--since", since, "--style", "json"]
        log_out = subprocess.run(cmd, capture_output=True, text=True, timeout=20)
        if log_out.returncode == 0 and log_out.stdout.strip():
            try:
                log_data = json.loads(log_out.stdout)
                events = log_data.get("predicate", {}).get("events", []) if isinstance(log_data, dict) else []
                for entry in events[:20]:
                    result["Events"].append({
                        "Time": entry.get("timestamp", ""),
                        "Subsystem": entry.get("subsystem", ""),
                        "Message": entry.get("eventMessage", "")[:200]
                    })
            except json.JSONDecodeError:
                pass

        result["TotalEvents"] = len(result["Events"]) + len(result["CrashReports"]) + len(result["KernelPanics"])
        if result["KernelPanics"] or result["CrashReports"] or result["Events"]:
            result["Status"] = "WARNING"

    except Exception as e:
        result["Status"] = "ERROR"
        result["Error"] = str(e)
        result["Note"] = "Unified logging requires macOS 10.12+ and may fail on older systems."

    return result

# ============================================================================
# PENDING REBOOT MODULE
# ============================================================================
def run_pending_reboot():
    """Check if a reboot is pending (e.g. from pending software updates)."""
    logger.info("Checking for pending reboot...")
    result = {"Status": "PASS", "RebootRequired": False, "UptimeDays": 0, "UptimeWarning": False}

    try:
        boot_time = int(subprocess.run(["sysctl", "-n", "kern.boottime"], capture_output=True, text=True, check=True).stdout.strip().split("{")[1].split("}")[0].split("sec = ")[1].split(",")[0].strip())
        uptime_seconds = int(time.time()) - boot_time
        uptime_days = round(uptime_seconds / 86400, 1)
        result["UptimeDays"] = uptime_days
        if uptime_days > 30:
            result["UptimeWarning"] = True
            result["Status"] = "WARNING"

        # Check for pending update reboot
        try:
            pending = subprocess.run(["defaults", "read", "/Library/Preferences/com.apple.SoftwareUpdate", "RecommendedUpdates"], capture_output=True, text=True)
            if pending.returncode == 0 and pending.stdout.strip():
                result["RebootRequired"] = result.get("RebootRequired", False) or True
        except Exception:
            pass

        # Check for flag files indicating reboot needed.
        # Note: /var/db/.AppleSetupDone is normal on configured Macs and is NOT a reboot flag.
        for flag in ["/var/db/receipts/.pkg.rebootRequired"]:
            if os.path.exists(flag):
                result["RebootRequired"] = True
                result.setdefault("RebootReasons", []).append(flag)

    except Exception as e:
        result["Status"] = "ERROR"
        result["Error"] = str(e)

    return result


# ============================================================================
# FIREWALL MODULE
# ============================================================================
def run_firewall_check():
    """Check macOS application firewall status."""
    logger.info("Checking firewall status...")
    result = {"Status": "PASS", "Enabled": False}

    try:
        # Prefer the supported socketfilterfw helper; fall back to preferences for older systems.
        sfw = subprocess.run(["/usr/libexec/ApplicationFirewall/socketfilterfw", "--getglobalstate"], capture_output=True, text=True)
        sfw_out = (sfw.stdout + sfw.stderr).strip()
        if sfw.returncode == 0 and sfw_out:
            result["RawState"] = sfw_out
            if "enabled" in sfw_out.lower():
                result["Enabled"] = True
            elif "disabled" in sfw_out.lower():
                result["Enabled"] = False
                result["Status"] = "WARNING"
        else:
            fw = subprocess.run(["defaults", "read", "/Library/Preferences/com.apple.alf", "globalstate"], capture_output=True, text=True)
            if fw.returncode == 0:
                state = fw.stdout.strip()
                if state == "1":
                    result["Enabled"] = True
                    result["Mode"] = "Allow essential services"
                elif state == "2":
                    result["Enabled"] = True
                    result["Mode"] = "Block all incoming connections"
                elif state == "0":
                    result["Enabled"] = False
                    result["Status"] = "WARNING"
                else:
                    result["Enabled"] = state
            else:
                result["Status"] = "WARNING"
                result["Note"] = "Could not determine firewall status without elevated access."

        stealth = subprocess.run(["/usr/libexec/ApplicationFirewall/socketfilterfw", "--getstealthmode"], capture_output=True, text=True)
        if stealth.returncode == 0 and stealth.stdout.strip():
            result["StealthMode"] = stealth.stdout.strip()
    except Exception as e:
        result["Status"] = "WARNING"
        result["Error"] = str(e)

    return result

# ============================================================================
# USER ACCOUNTS MODULE
# ============================================================================
def run_user_accounts():
    """Audit user accounts."""
    logger.info("Auditing user accounts...")
    result = {"Status": "PASS", "Accounts": [], "AdminCount": 0, "EnabledCount": 0, "FlaggedCount": 0, "FlaggedAccounts": []}

    try:
        # Get list of local users
        users = subprocess.run(["dscl", ".", "list", "/Users"], capture_output=True, text=True, check=True).stdout.strip().split('\n')
        
        for user in users:
            user = user.strip()
            if user.startswith('_') or user in ('daemon', 'nobody', 'root', 'Guest'):
                continue  # System accounts
            
            try:
                user_info = subprocess.run(["dscl", ".", "read", f"/Users/{user}"], capture_output=True, text=True)
                info = user_info.stdout
                
                is_admin = False
                try:
                    admin_check = subprocess.run(["dseditgroup", "-o", "checkmember", "-m", user, "admin"], capture_output=True, text=True)
                    is_admin = admin_check.returncode == 0 and "yes" in (admin_check.stdout + admin_check.stderr).lower()
                except Exception:
                    is_admin = "admin" in info.lower()
                if is_admin:
                    result["AdminCount"] += 1
                result["EnabledCount"] += 1

                uid_match = re.search(r"UniqueID:\s*(\d+)", info)
                acct = {"Name": user, "IsAdmin": is_admin, "UID": uid_match.group(1) if uid_match else "?"}
                
                # Check last login (requires last command)
                try:
                    last = subprocess.run(["last", user], capture_output=True, text=True)
                    if last.stdout.strip():
                        lines = last.stdout.strip().split('\n')
                        first = lines[0]
                        acct["LastLogin"] = first[:50]
                except Exception:
                    acct["LastLogin"] = "?"
                
                result["Accounts"].append(acct)
                
                # Flag if admin and no last login info (suspicious but not actionable)
                if is_admin and acct.get("LastLogin", "?") == "?":
                    pass  # Not necessarily a concern for Mac

            except Exception:
                pass

    except Exception as e:
        result["Status"] = "ERROR"
        result["Error"] = str(e)

    try:
        guest = subprocess.run(["defaults", "read", "/Library/Preferences/com.apple.loginwindow", "GuestEnabled"], capture_output=True, text=True)
        result["GuestEnabled"] = guest.returncode == 0 and guest.stdout.strip() == "1"
        if result["GuestEnabled"] and result.get("Status") != "ERROR":
            result["Status"] = "WARNING"
    except Exception:
        pass

    try:
        auto = subprocess.run(["defaults", "read", "/Library/Preferences/com.apple.loginwindow", "autoLoginUser"], capture_output=True, text=True)
        if auto.returncode == 0 and auto.stdout.strip():
            result["AutoLoginUser"] = auto.stdout.strip()
            if result.get("Status") != "ERROR":
                result["Status"] = "WARNING"
    except Exception:
        pass

    return result


# ============================================================================
# TEMP FILES MODULE
# ============================================================================
def run_temp_files():
    """Check temp and cache directories."""
    logger.info("Checking temp and cache files...")
    result = {"Status": "PASS", "Directories": []}

    dirs_to_check = [
        ("/tmp", "System temp"),
        ("/var/tmp", "Var temp"),
        (str(Path.home() / "Library/Caches"), "User caches"),
    ]

    # Browser caches
    for browser in ["Google/Chrome", "Safari", "Firefox"]:
        cache_dir = Path.home() / "Library/Caches" / browser
        if cache_dir.exists():
            dirs_to_check.append((str(cache_dir), f"{browser} cache"))

    for dpath, label in dirs_to_check:
        try:
            if os.path.exists(dpath):
                total_size = 0
                file_count = 0
                for root, dirs, files in os.walk(dpath):
                    for f in files:
                        try:
                            fp = os.path.join(root, f)
                            total_size += os.path.getsize(fp)
                            file_count += 1
                        except OSError:
                            pass
                size_mb = round(total_size / (1024 * 1024), 1)
                warning = size_mb > 1024  # > 1 GB
                result["Directories"].append({
                    "Path": dpath,
                    "Label": label,
                    "SizeMB": size_mb,
                    "FileCount": file_count,
                    "Warning": warning
                })
                if warning and result["Status"] != "ERROR":
                    result["Status"] = "WARNING"
        except Exception:
            pass

    return result


# ============================================================================
# STARTUP / LOGIN ITEMS MODULE
# ============================================================================
def run_startup_check():
    """Check login items and LaunchAgents."""
    logger.info("Checking startup/login items...")
    result = {"Status": "PASS", "LoginItems": [], "LaunchAgents": [], "LaunchDaemons": []}

    # Login items (modern macOS)
    try:
        # Use osascript to query login items (macOS 13+)
        script = 'tell application "System Events" to get the name of every login item'
        login = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if login.returncode == 0 and login.stdout.strip():
            for item in login.stdout.strip().split(', '):
                result["LoginItems"].append({"Name": item.strip()})
    except Exception:
        pass

    # LaunchAgents (user)
    user_agents = Path.home() / "Library/LaunchAgents"
    if user_agents.exists():
        for f in user_agents.iterdir():
            if f.suffix == ".plist":
                result["LaunchAgents"].append({"Name": f.name, "Path": str(f)})

    # LaunchDaemons (system)
    sys_daemons = Path("/Library/LaunchDaemons")
    if sys_daemons.exists():
        for f in sys_daemons.iterdir():
            if f.suffix == ".plist":
                result["LaunchDaemons"].append({"Name": f.name, "Path": str(f)})

    return result


# ============================================================================
# SCHEDULED TASKS MODULE
# ============================================================================
def run_scheduled_tasks():
    """Check LaunchAgent/LaunchDaemon schedules."""
    logger.info("Checking scheduled tasks...")
    result = {"Status": "PASS", "Tasks": []}

    # Only check known plist directories – much faster than `launchctl list`
    # which enumerates hundreds of system services
    search_dirs = [
        Path.home() / "Library/LaunchAgents",
        Path("/Library/LaunchAgents"),
        Path("/Library/LaunchDaemons"),
    ]

    seen = set()
    for d in search_dirs:
        if d.is_dir():
            for plist in d.glob("*.plist"):
                try:
                    label = plist.stem
                    if label in seen:
                        continue
                    seen.add(label)
                    # Quick check: is this our maintenance agent?
                    is_ours = "pcmasterclass" in label.lower()
                    result["Tasks"].append({
                        "Label": label,
                        "Status": "External" if not is_ours else "Ours",
                        "Path": str(plist)
                    })
                except Exception:
                    pass

    # Also verify our agent is actually loaded
    ours_found = any("pcmasterclass" in t["Label"].lower() for t in result["Tasks"])
    if not ours_found:
        result["Tasks"].append({
            "Label": "com.pcmasterclass.maintenance.agent (not found in search dirs)",
            "Status": "NOT_INSTALLED",
            "Path": ""
        })

    return result


# ============================================================================
# BROWSER EXTENSIONS MODULE
# ============================================================================
def run_browser_extensions():
    """Check common Chromium and Firefox browser extensions across all profiles."""
    logger.info("Checking browser extensions...")
    result = {"Status": "PASS", "Extensions": [], "ChromeExtensions": [], "SafariExtensions": []}

    chromium_browsers = [
        ("Chrome", Path.home() / "Library/Application Support/Google/Chrome"),
        ("Edge", Path.home() / "Library/Application Support/Microsoft Edge"),
        ("Brave", Path.home() / "Library/Application Support/BraveSoftware/Brave-Browser"),
    ]

    def add_extension(browser, profile, ext_id, name, version=""):
        item = {"Browser": browser, "Profile": profile, "ID": ext_id, "Name": name, "Version": version}
        result["Extensions"].append(item)
        if browser == "Chrome":
            result["ChromeExtensions"].append(item)

    for browser, base in chromium_browsers:
        if not base.exists():
            continue
        for profile in base.iterdir():
            ext_dir = profile / "Extensions"
            if not ext_dir.is_dir():
                continue
            for ext_id in ext_dir.iterdir():
                if not ext_id.is_dir():
                    continue
                try:
                    versions = [v for v in ext_id.iterdir() if v.is_dir()]
                    if not versions:
                        continue
                    latest = sorted(versions)[-1]
                    manifest = latest / "manifest.json"
                    if manifest.exists():
                        data = json.loads(manifest.read_text(errors="ignore"))
                        name = data.get("name", ext_id.name)
                        version = data.get("version", latest.name)
                        add_extension(browser, profile.name, ext_id.name, name, version)
                except Exception:
                    pass

    firefox_base = Path.home() / "Library/Application Support/Firefox/Profiles"
    if firefox_base.exists():
        for profile in firefox_base.iterdir():
            ext_json = profile / "extensions.json"
            if not ext_json.exists():
                continue
            try:
                data = json.loads(ext_json.read_text(errors="ignore"))
                for addon in data.get("addons", []):
                    if addon.get("type") == "extension":
                        add_extension("Firefox", profile.name, addon.get("id", ""), addon.get("defaultLocale", {}).get("name", addon.get("id", "")), addon.get("version", ""))
            except Exception:
                pass

    safari_dirs = [Path.home() / "Library/Safari/Extensions", Path.home() / "Library/Containers/com.apple.Safari/Data/Library/Safari/Extensions"]
    for sdir in safari_dirs:
        if sdir.exists():
            for item in sdir.iterdir():
                if item.is_file() or item.is_dir():
                    result["SafariExtensions"].append({"Name": item.name, "Path": str(item)})

    return result

# ============================================================================
# SERVICE STATUS MODULE
# ============================================================================
def run_service_status():
    """Check support-relevant third-party agents and key macOS services."""
    logger.info("Checking service status...")
    result = {"Status": "PASS", "Services": []}

    service_patterns = [
        ("TeamViewer", ["TeamViewer"]),
        ("Malwarebytes", ["Malwarebytes", "MBAM"]),
        ("iDrive", ["iDrive", "IDrive"]),
        ("Dropbox", ["Dropbox"]),
        ("OneDrive", ["OneDrive"]),
        ("Google Drive", ["Google Drive", "DriveFS"]),
        ("Tailscale", ["Tailscale"]),
        ("WireGuard", ["WireGuard"]),
        ("OpenVPN", ["OpenVPN"]),
        ("Cisco AnyConnect", ["AnyConnect", "Cisco Secure Client"]),
        ("FortiClient", ["FortiClient"]),
        ("CoreAudio", ["coreaudiod"]),
        ("Spotlight", ["mds"]),
    ]

    ps = subprocess.run(["ps", "-axo", "pid=,comm="], capture_output=True, text=True)
    processes = ps.stdout if ps.returncode == 0 else ""
    lower_processes = processes.lower()

    for name, patterns in service_patterns:
        running = any(pattern.lower() in lower_processes for pattern in patterns)
        app_paths = [Path("/Applications") / f"{name}.app", Path.home() / "Applications" / f"{name}.app"]
        installed = any(path.exists() for path in app_paths) or running
        result["Services"].append({"Name": name, "Installed": installed, "Running": running})

    return result

# ============================================================================
# NETWORK CONFIG MODULE
# ============================================================================
def run_network_config():
    """Check network configuration."""
    logger.info("Checking network configuration...")
    result = {"Status": "PASS", "Interfaces": [], "WiFi": "", "DNS": []}

    try:
        # Get active interfaces
        route = subprocess.run(["netstat", "-rn"], capture_output=True, text=True, check=True).stdout
        for line in route.split('\n'):
            if line.startswith('default') :
                parts = line.split()
                if len(parts) >= 2:
                    result["DefaultGateway"] = parts[1]
                    break

        # DNS
        dns = subprocess.run(["scutil", "--dns"], capture_output=True, text=True, check=True).stdout
        for line in dns.split('\n'):
            if "nameserver" in line:
                ip = line.split()[-1].strip()
                if ip not in result["DNS"]:
                    result["DNS"].append(ip)

        # WiFi
        try:
            wifi = subprocess.run(["/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport", "-I"], capture_output=True, text=True)
            if wifi.returncode == 0:
                for line in wifi.stdout.split('\n'):
                    if " SSID" in line:
                        result["WiFi"] = line.split(":")[-1].strip()
        except Exception:
            pass

    except Exception as e:
        result["Status"] = "WARNING"
        result["Error"] = str(e)

    return result



# ============================================================================
# MAINTENANCE AGENT HEALTH MODULE
# ============================================================================
def run_agent_health(email_to=""):
    """Verify the PC Master Class maintenance agent installation and update path."""
    logger.info("Checking PC Master Class maintenance agent health...")
    result = {"Status": "PASS", "Checks": []}

    def add_check(name, ok, detail=""):
        result["Checks"].append({"Name": name, "OK": bool(ok), "Detail": detail})
        if not ok and result["Status"] != "ERROR":
            result["Status"] = "WARNING"

    script_path = Path(sys.argv[0]).expanduser().resolve()
    expected_script = Path.home() / "Library/PCMasterClass/pcm_mac_maintenance.py"
    plist_path = Path.home() / "Library/LaunchAgents/com.pcmasterclass.maintenance.agent.plist"
    report_dir = Path.home() / "Library/PCMasterClass/Reports"

    add_check("Script exists", script_path.exists(), str(script_path))
    add_check("Installed in expected location", script_path == expected_script or expected_script.exists(), str(expected_script))
    add_check("Reports directory exists", report_dir.exists(), str(report_dir))
    add_check("LaunchAgent plist exists", plist_path.exists(), str(plist_path))

    if plist_path.exists():
        try:
            with plist_path.open("rb") as f:
                plist = plistlib.load(f)
            result["LaunchAgentLabel"] = plist.get("Label", "")
            result["LaunchAgentSchedule"] = plist.get("StartCalendarInterval", "")
            args = plist.get("ProgramArguments", [])
            add_check("LaunchAgent points at maintenance script", any("pcm_mac_maintenance.py" in str(a) for a in args), " ".join(map(str, args)))
        except Exception as e:
            add_check("LaunchAgent plist readable", False, str(e))

    loaded = subprocess.run(["launchctl", "list", "com.pcmasterclass.maintenance.agent"], capture_output=True, text=True)
    add_check("LaunchAgent loaded", loaded.returncode == 0, loaded.stdout.strip() or loaded.stderr.strip())

    if email_to:
        cred = load_smtp_credentials(email_to)
        add_check("SMTP credential available", bool(cred), email_to or "reports@pcmasterclass.com.au")

    try:
        req = urllib.request.Request(UPDATE_URL, headers={})
        with urllib.request.urlopen(req, timeout=10) as response:
            remote_content = response.read().decode("utf-8")
        match = re.search(r'SCRIPT_VERSION\s*=\s*"([^"]+)"', remote_content)
        result["RemoteVersion"] = match.group(1) if match else "Unknown"
        add_check("Self-update URL reachable", True, UPDATE_URL)
    except Exception as e:
        add_check("Self-update URL reachable", False, f"{UPDATE_URL} — {e}")

    return result

# ============================================================================
# HTML REPORT GENERATOR
# ============================================================================
def generate_html_report(results, client_name="", computer_name="", script_version=SCRIPT_VERSION):
    """Generate a branded HTML maintenance report matching the PC version style."""
    
    hostname = platform.node()
    os_ver = platform.mac_ver()[0]
    sys_info = results.get("SystemInfo", {})
    display_computer_name = build_hardware_display_name(sys_info, computer_name)
    now = datetime.now().strftime("%A, %d %B %Y %I:%M %p")
    
    # Count warnings and errors
    warning_count = sum(1 for r in results.values() if isinstance(r, dict) and r.get("Status") == "WARNING")
    error_count = sum(1 for r in results.values() if isinstance(r, dict) and r.get("Status") == "ERROR")
    
    if error_count > 0:
        overall_status = "Errors Found"
        overall_color = "#dc2626"
    elif warning_count > 0:
        overall_status = "Warnings Found"
        overall_color = "#d97706"
    else:
        overall_status = "All Clear"
        overall_color = "#16a34a"
    overall_badge = f'<span style="background:{overall_color};color:white;padding:4px 12px;border-radius:4px;font-weight:bold;font-size:0.9em;">{overall_status}</span>'
    
    def status_badge(status):
        colors = {"PASS": "#16a34a", "CLEAN": "#16a34a", "UP TO DATE": "#16a34a", "REPAIRED": "#16a34a",
                  "WARNING": "#d97706", "INFO": "#3b82f6", "ERROR": "#dc2626", "SKIPPED": "#94a3b8"}
        color = colors.get(status, "#94a3b8")
        return f'<span style="background:{color};color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;font-weight:bold;">{status}</span>'
    
    def heading_class(status):
        classes = {"PASS": "heading-ok", "CLEAN": "heading-ok", "UP TO DATE": "heading-ok",
                   "WARNING": "heading-warn", "ERROR": "heading-error"}
        return classes.get(status, "heading-info")
    
    # Build summary items
    summary_items = [
        ("disk", "Disk Health", results.get("DiskHealth", {}).get("Status", "INFO")),
        ("updates", "Software Updates", results.get("SoftwareUpdates", {}).get("Status", "INFO")),
        ("backup", "Backup Status", results.get("Backup", {}).get("Status", "INFO")),
        ("malwarebytes", "Malwarebytes", results.get("Malwarebytes", {}).get("Status", "INFO")),
        ("security", "macOS Security", results.get("Security", {}).get("Status", "INFO")),
        ("logs", "System Log Errors", results.get("LogErrors", {}).get("Status", "INFO")),
        ("reboot", "Pending Reboot", results.get("PendingReboot", {}).get("Status", "INFO")),
        ("firewall", "Firewall", results.get("Firewall", {}).get("Status", "INFO")),
        ("users", "User Accounts", results.get("UserAccounts", {}).get("Status", "INFO")),
        ("temp", "Temp/Cache Files", results.get("TempFiles", {}).get("Status", "INFO")),
        ("startup", "Startup/Login Items", results.get("Startup", {}).get("Status", "INFO")),
        ("tasks", "Scheduled Tasks", results.get("ScheduledTasks", {}).get("Status", "INFO")),
        ("extensions", "Browser Extensions", results.get("BrowserExtensions", {}).get("Status", "INFO")),
        ("services", "Service Status", results.get("ServiceStatus", {}).get("Status", "INFO")),
        ("network", "Network Config", results.get("NetworkConfig", {}).get("Status", "INFO")),
        ("agent", "Maintenance Agent", results.get("AgentHealth", {}).get("Status", "INFO")),
    ]
    
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Maintenance Report — {display_computer_name}</title>
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f1f5f9; color: #1e293b; line-height: 1.5; padding: 20px; }}
    .container {{ max-width: 900px; margin: 0 auto; }}
    .header {{ background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%); color: white; padding: 30px; border-radius: 12px 12px 0 0; }}
    .header h1 {{ font-size: 1.5em; margin-bottom: 5px; }}
    .header .subtitle {{ opacity: 0.8; font-size: 0.9em; }}
    .header .overall {{ margin-top: 15px; font-size: 1.1em; }}
    .section {{ background: white; border: 1px solid #e2e8f0; margin-bottom: 2px; padding: 20px 25px; }}
    .section:last-child {{ border-radius: 0 0 12px 12px; margin-bottom: 20px; }}
    .section h2 {{ font-size: 1.1em; color: #334155; margin-bottom: 10px; display: flex; align-items: center; gap: 10px; }}
    .section h2 .icon {{ font-size: 1.3em; }}
    table {{ width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 0.9em; }}
    th {{ background: #f8fafc; text-align: left; padding: 8px 12px; border-bottom: 2px solid #e2e8f0; color: #64748b; font-weight: 600; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }}
    tr:hover td {{ background: #f8fafc; }}
    .flagged {{ background: #fef2f2 !important; }}
    .detail {{ color: #64748b; font-size: 0.85em; }}
    .warning-text {{ color: #d97706; font-weight: 600; }}
    .error-text {{ color: #dc2626; font-weight: 600; }}
    .footer {{ text-align: center; color: #94a3b8; font-size: 0.8em; padding: 15px; }}
    html {{ scroll-behavior: smooth; }}
    .heading-ok {{ color: #16a34a; }}
    .heading-warn {{ color: #d97706; }}
    .heading-error {{ color: #dc2626; }}
    .heading-info {{ color: #334155; }}
    details summary {{ cursor: pointer; list-style: none; display: flex; align-items: center; gap: 10px; font-size: 1.1em; color: #334155; font-weight: bold; padding: 2px 0; }}
    details summary::-webkit-details-marker {{ display: none; }}
    details summary::before {{ content: '\\25B6'; font-size: 0.7em; transition: transform 0.2s ease; display: inline-block; min-width: 14px; }}
    details[open] summary::before {{ transform: rotate(90deg); }}
    details .section-body {{ padding-top: 10px; }}
</style>
</head>
<body>
<div class="container">

<div class="header">
    <h1>PC Master Class — Maintenance Report</h1>
    <div class="subtitle">{client_name + ' • ' if client_name else ''}{display_computer_name} • macOS {os_ver} • {now}</div>
    <div class="overall">Overall Status: {overall_badge} &nbsp; ({warning_count} warning(s), {error_count} error(s))</div>
</div>

<!-- CHECK RESULTS SUMMARY -->
<div class="section" style="padding:15px 25px;">
    <h2 style="margin-bottom:8px;"><span class="icon">&#x1F4CB;</span> Check Results Summary</h2>
    <table style="font-size:0.88em;">
"""
    
    col = 0
    for anchor, label, status in summary_items:
        if col % 2 == 0:
            html += "<tr>"
        html += f"<td style='padding:4px 10px;width:50%;'><a href='#{anchor}' style='text-decoration:none;color:#1e40af;'>{label}</a></td><td style='padding:4px 6px;'>{status_badge(status)}</td>"
        if col % 2 == 1:
            html += "</tr>"
        col += 1
    if col % 2 == 1:
        html += "<td></td><td></td></tr>"
    html += "</table></div>"
    
    # SYSTEM INFO
    sys = results.get("SystemInfo", {})
    html += f"""
<div class="section">
    <h2><span class="icon">&#x1F4BB;</span> System Information</h2>
    <table>
        {"<tr><td><strong>Client</strong></td><td><strong style='font-size:1.1em;'>" + client_name + "</strong></td></tr>" if client_name else ""}
        <tr><td><strong>Computer</strong></td><td><strong style='font-size:1.1em;'>{display_computer_name}</strong></td></tr>
        <tr><td><strong>Hostname</strong></td><td>{sys.get("Hostname", hostname)}</td></tr>
        <tr><td><strong>Make / Model</strong></td><td>{sys.get("Manufacturer", "Apple")} {sys.get("ModelName", "Unknown")} ({sys.get("ModelIdentifier", sys.get("Model", ""))})</td></tr>
        <tr><td><strong>Processor</strong></td><td>{sys.get("ProcessorName", sys.get("CPU", "Unknown"))}</td></tr>
        <tr><td><strong>Serial Number</strong></td><td>{sys.get("SerialNumber", "Unknown")}</td></tr>
        <tr><td><strong>macOS</strong></td><td>{sys.get("OS", os_ver)} ({sys.get("Build", "")})</td></tr>
        <tr><td><strong>CPU Detail</strong></td><td>{sys.get("CPU", "Unknown")}</td></tr>
        <tr><td><strong>RAM</strong></td><td>{sys.get("RAM_GB", "?")} GB</td></tr>
        <tr><td><strong>FileVault</strong></td><td>{sys.get("FileVault", "Unknown")}</td></tr>
        {"<tr><td><strong>Battery</strong></td><td>" + str(sys.get("Battery", "")) + "</td></tr>" if sys.get("Battery") else ""}
        {"<tr><td><strong>Battery Health</strong></td><td>" + str(sys.get("BatteryCondition", "Unknown")) + " — " + str(sys.get("BatteryCycleCount", "?")) + " cycles</td></tr>" if sys.get("BatteryCondition") or sys.get("BatteryCycleCount") else ""}
        <tr><td><strong>Last Boot</strong></td><td>{sys.get("LastBoot", "?")} (Uptime: {sys.get("UptimeDays", "?")} days)</td></tr>
        <tr><td><strong>Script Version</strong></td><td>v{script_version}</td></tr>
    </table>
</div>
"""
    
    # DISK HEALTH
    disk = results.get("DiskHealth", {})
    disk_open = "" if disk.get("Status") in ("PASS", "CLEAN") else " open"
    disk_class = heading_class(disk.get("Status", "INFO"))
    html += f"""
<div class="section" id="disk">
<details{disk_open}><summary class="{disk_class}"><span class="icon">&#x1F4BE;</span> Disk Health {status_badge(disk.get("Status", "INFO"))}</summary>
<div class="section-body">
{"<p class='error-text'>Error: " + disk.get("Error", "") + "</p>" if disk.get("Status") == "ERROR" else ""}
<table><tr><th>Disk</th><th>Type</th><th>Size</th><th>Health</th></tr>
"""
    for d in disk.get("Disks", []):
        html += f"<tr><td>{d.get('Name', '')}</td><td>{d.get('Type', '')}</td><td>{d.get('SizeGB', '')} GB</td><td>{d.get('SMART', 'N/A')}</td></tr>"
    html += "</table>"
    
    if disk.get("Volumes"):
        html += '<table><tr><th>Filesystem</th><th>Mount</th><th>Size</th><th>Used</th><th>Avail</th><th>Used %</th></tr>'
        for v in disk["Volumes"]:
            cls = " class='flagged'" if v.get("Warning") else ""
            pct_text = f"<span class='warning-text'>{v.get('UsedPercent', 0)}% WARNING</span>" if v.get("Warning") else f"{v.get('UsedPercent', 0)}%"
            html += f"<tr{cls}><td>{v.get('Filesystem', '')}</td><td>{v.get('Mount', '')}</td><td>{v.get('Size', '')}</td><td>{v.get('Used', '')}</td><td>{v.get('Avail', '')}</td><td>{pct_text}</td></tr>"
        html += "</table>"
    if disk.get("SMARTNote"):
        html += f"<p class='detail'>{disk['SMARTNote']}</p>"
    html += "</div></details></div>"

    # SOFTWARE UPDATES
    upd = results.get("SoftwareUpdates", {})
    upd_open = "" if upd.get("Status") in ("PASS", "CLEAN") else " open"
    upd_class = heading_class(upd.get("Status", "INFO"))
    html += f"""
<div class="section" id="updates">
<details{upd_open}><summary class="{upd_class}"><span class="icon">&#x1F504;</span> Software Updates {status_badge(upd.get("Status", "INFO"))}</summary>
<div class="section-body">
"""
    if upd.get("Status") == "SKIPPED":
        html += "<p class='detail'>Skipped by configuration.</p>"
    elif upd.get("PendingCount", 0) == 0:
        html += "<p>No pending updates. System is up to date.</p>"
    else:
        html += f"<p>{upd.get('PendingCount', 0)} macOS update(s) available.</p><table><tr><th>Update</th></tr>"
        for u in upd.get("Updates", []):
            html += f"<tr><td>{u.get('Title', '')}</td></tr>"
        html += "</table>"
    if upd.get("HomebrewUpdates"):
        html += f"<p class='detail'>{upd['HomebrewUpdates']} Homebrew package(s) outdated.</p>"
    if upd.get("RebootRequired"):
        html += "<p class='warning-text'>A restart is required to complete pending updates.</p>"
    html += "</div></details></div>"

    # BACKUP
    bak = results.get("Backup", {})
    bak_open = "" if bak.get("Status") in ("PASS", "CLEAN") else " open"
    bak_class = heading_class(bak.get("Status", "INFO"))
    html += f"""
<div class="section" id="backup">
<details{bak_open}><summary class="{bak_class}"><span class="icon">&#x2601;</span> Backup Status {status_badge(bak.get("Status", "INFO"))}</summary>
<div class="section-body">
"""
    if bak.get("iDrive", {}).get("Installed"):
        idr = bak["iDrive"]
        html += f"""
    <h3>iDrive</h3>
    <table>
        <tr><td><strong>Running</strong></td><td>{'Yes' if idr.get('Running') else '<span class="warning-text">No</span>'}</td></tr>
        {"<tr><td><strong>Last Log</strong></td><td>" + idr.get('LastLog', '?') + "</td></tr>" if idr.get('LastLog') else ""}
    </table>
"""
    else:
        html += "<p class='detail'>iDrive not installed.</p>"

    if bak.get("TimeMachine", {}).get("Configured"):
        tm = bak["TimeMachine"]
        html += f"""
    <h3>Time Machine</h3>
    <table>
        <tr><td><strong>Last Backup</strong></td><td>{tm.get('LastBackupDate', 'Unknown')}</td></tr>
        {"<tr><td><strong>Running Now</strong></td><td>" + ("Yes" if tm.get("Running") else "No") + "</td></tr>" if "Running" in tm else ""}
        {"<tr><td><strong>Local Snapshots</strong></td><td>" + str(tm.get("LocalSnapshotCount")) + "</td></tr>" if "LocalSnapshotCount" in tm else ""}
        {"<tr><td><strong>Days Since</strong></td><td><span class='warning-text'>" + str(tm.get('DaysSince', '?')) + " days</span></td></tr>" if tm.get('DaysSince') else ""}
        <tr><td><strong>Destinations</strong></td><td><pre>{tm.get('Destinations', 'N/A')}</pre></td></tr>
    </table>
"""
    else:
        html += "<p class='detail'>Time Machine not configured.</p>"

    if bak.get("Status") == "WARNING":
        html += "<p class='warning-text'>No backup solution detected or backup is stale.</p>"
    html += "</div></details></div>"

    # MALWAREBYTES
    mb = results.get("Malwarebytes", {})
    mb_open = "" if mb.get("Status") in ("PASS", "CLEAN") else " open"
    mb_class = heading_class(mb.get("Status", "INFO"))
    html += f"""
<div class="section" id="malwarebytes">
<details{mb_open}><summary class="{mb_class}"><span class="icon">&#x1F6E1;</span> Malwarebytes Endpoint Protection {status_badge(mb.get("Status", "INFO"))}</summary>
<div class="section-body">
"""
    if mb.get("Installed"):
        html += f"""
    <table>
        <tr><td><strong>Product</strong></td><td>{mb.get('ProductType', 'Unknown')}</td></tr>
        <tr><td><strong>Service Running</strong></td><td>{'Yes' if mb.get('ServiceRunning') else '<span class="warning-text">No</span>'}</td></tr>
    </table>
"""
    else:
        html += f"<p class='detail'>{mb.get('Note', 'Malwarebytes not installed.')}</p>"
    html += "</div></details></div>"

    # MACOS SECURITY (replaces Windows Defender)
    sec = results.get("Security", {})
    sec_open = "" if sec.get("Status") in ("PASS", "CLEAN") else " open"
    sec_class = heading_class(sec.get("Status", "INFO"))
    html += f"""
<div class="section" id="security">
<details{sec_open}><summary class="{sec_class}"><span class="icon">&#x1F6E1;</span> macOS Security {status_badge(sec.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table>
        <tr><td><strong>Gatekeeper</strong></td><td>{sec.get('Gatekeeper', 'Unknown')}</td></tr>
        <tr><td><strong>XProtect</strong></td><td>{sec.get('XProtect', 'Unknown')}</td></tr>
        {"<tr><td><strong>XProtect Install Date</strong></td><td>" + str(sec.get('XProtectInstallDate', '')) + "</td></tr>" if sec.get('XProtectInstallDate') else ""}
        <tr><td><strong>System Integrity Protection</strong></td><td>{sec.get('SIP', 'Unknown')}</td></tr>
    </table>
</div></details></div>
"""

    # LOG ERRORS
    ev = results.get("LogErrors", {})
    ev_open = "" if ev.get("Status") in ("PASS", "CLEAN") else " open"
    ev_class = heading_class(ev.get("Status", "INFO"))
    html += f"""
<div class="section" id="logs">
<details{ev_open}><summary class="{ev_class}"><span class="icon">&#x1F4CB;</span> Actionable Logs / Crashes (Last {ev.get('HoursChecked', 24)}h) {status_badge(ev.get('Status', 'INFO'))}</summary>
<div class="section-body">
"""
    if ev.get("TotalEvents", 0) == 0:
        html += f"<p>No critical or error events in the last {ev.get('HoursChecked', 24)} hours. All clear.</p>"
    else:
        html += f"<p>{ev.get('TotalEvents')} event(s) found.</p><table><tr><th>Time</th><th>Subsystem</th><th>Message</th></tr>"
        for e in ev.get("Events", [])[:20]:
            html += f"<tr><td style='white-space:nowrap'>{e.get('Time', '?')}</td><td>{e.get('Subsystem', '?')}</td><td>{e.get('Message', '')[:100]}</td></tr>"
        html += "</table>"

    if ev.get("CrashReports"):
        html += f"<h3 style='color:#d97706;margin:10px 0 5px;font-size:0.95em;'>Recent Crash Reports</h3><ul>"
        for c in ev["CrashReports"]:
            html += f"<li>{c}</li>"
        html += "</ul>"
    html += "</div></details></div>"

    # PENDING REBOOT
    reb = results.get("PendingReboot", {})
    reb_open = "" if reb.get("Status") in ("PASS", "CLEAN") else " open"
    reb_class = heading_class(reb.get("Status", "INFO"))
    html += f"""
<div class="section" id="reboot">
<details{reb_open}><summary class="{reb_class}"><span class="icon">&#x1F504;</span> Pending Reboot {status_badge(reb.get("Status", "INFO"))}</summary>
<div class="section-body">
"""
    if reb.get("RebootRequired"):
        html += "<p class='warning-text'>A restart is required to complete pending updates.</p>"
    elif reb.get("UptimeWarning"):
        html += f"<p class='warning-text'>System has been running for {reb.get('UptimeDays', '?')} days without a restart. Consider scheduling a restart.</p>"
    else:
        html += f"<p>No reboot pending. Uptime: {reb.get('UptimeDays', '?')} days.</p>"
    html += "</div></details></div>"

    # FIREWALL
    fw = results.get("Firewall", {})
    fw_open = "" if fw.get("Status") in ("PASS", "CLEAN") else " open"
    fw_class = heading_class(fw.get("Status", "INFO"))
    html += f"""
<div class="section" id="firewall">
<details{fw_open}><summary class="{fw_class}"><span class="icon">&#x1F6E1;</span> Firewall {status_badge(fw.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table>
        <tr><td><strong>Enabled</strong></td><td>{'Yes' if fw.get('Enabled') else '<span class="error-text">NO</span>'}</td></tr>
        {"<tr><td><strong>Mode</strong></td><td>" + fw.get('Mode', '') + "</td></tr>" if fw.get('Mode') else ""}
        {"<tr><td><strong>Stealth Mode</strong></td><td>" + fw.get('StealthMode', '') + "</td></tr>" if fw.get('StealthMode') else ""}
        {"<tr><td><strong>Note</strong></td><td>" + fw.get('Note', '') + "</td></tr>" if fw.get('Note') else ""}
    </table>
</div></details></div>
"""

    # USER ACCOUNTS
    ua = results.get("UserAccounts", {})
    ua_open = "" if ua.get("Status") in ("PASS", "CLEAN") else " open"
    ua_class = heading_class(ua.get("Status", "INFO"))
    html += f"""
<div class="section" id="users">
<details{ua_open}><summary class="{ua_class}"><span class="icon">&#x1F464;</span> User Accounts {status_badge(ua.get("Status", "INFO"))}</summary>
<div class="section-body">
    <p>{ua.get('EnabledCount', '?')} local account(s), {ua.get('AdminCount', '?')} with admin rights.</p>
    {"<p class='warning-text'>Guest account is enabled.</p>" if ua.get("GuestEnabled") else ""}
    {"<p class='warning-text'>Automatic login enabled for: " + str(ua.get("AutoLoginUser")) + "</p>" if ua.get("AutoLoginUser") else ""}
    <table><tr><th>Account</th><th>Admin</th><th>Last Login</th></tr>
"""
    for acct in ua.get("Accounts", []):
        html += f"<tr><td>{acct.get('Name', '')}</td><td>{'Yes' if acct.get('IsAdmin') else 'No'}</td><td>{acct.get('LastLogin', '?')}</td></tr>"
    html += "</table></div></details></div>"

    # TEMP FILES
    tf = results.get("TempFiles", {})
    tf_open = "" if tf.get("Status") in ("PASS", "CLEAN") else " open"
    tf_class = heading_class(tf.get("Status", "INFO"))
    html += f"""
<div class="section" id="temp">
<details{tf_open}><summary class="{tf_class}"><span class="icon">&#x1F4C1;</span> Temp/Cache Files {status_badge(tf.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table><tr><th>Directory</th><th>Label</th><th>Size</th><th>Files</th></tr>
"""
    for d in tf.get("Directories", []):
        cls = " class='flagged'" if d.get("Warning") else ""
        sz = f"<span class='warning-text'>{d.get('SizeMB', 0)} MB</span>" if d.get("Warning") else f"{d.get('SizeMB', 0)} MB"
        html += f"<tr{cls}><td>{d.get('Path', '')}</td><td>{d.get('Label', '')}</td><td>{sz}</td><td>{d.get('FileCount', 0)}</td></tr>"
    html += "</table></div></details></div>"

    # STARTUP / LOGIN ITEMS
    st = results.get("Startup", {})
    st_open = "" if st.get("Status") in ("PASS", "CLEAN") else " open"
    st_class = heading_class(st.get("Status", "INFO"))
    html += f"""
<div class="section" id="startup">
<details{st_open}><summary class="{st_class}"><span class="icon">&#x1F504;</span> Startup/Login Items {status_badge(st.get("Status", "INFO"))}</summary>
<div class="section-body">
    <h3 style="margin:10px 0 5px;font-size:0.95em;">Login Items</h3>
    <ul>
"""
    if st.get("LoginItems"):
        for li in st["LoginItems"]:
            html += f"<li>{li.get('Name', '')}</li>"
    else:
        html += "<li class='detail'>None found</li>"
    html += "</ul>"
    
    html += "<h3 style='margin:10px 0 5px;font-size:0.95em;'>LaunchAgents</h3><ul>"
    if st.get("LaunchAgents"):
        for la in st["LaunchAgents"][:10]:
            html += f"<li>{la.get('Name', '')}</li>"
        if len(st["LaunchAgents"]) > 10:
            html += f"<li class='detail'>... and {len(st['LaunchAgents']) - 10} more</li>"
    else:
        html += "<li class='detail'>None found</li>"
    html += "</ul></div></details></div>"

    # SCHEDULED TASKS
    sch = results.get("ScheduledTasks", {})
    sch_open = "" if sch.get("Status") in ("PASS", "CLEAN") else " open"
    sch_class = heading_class(sch.get("Status", "INFO"))
    html += f"""
<div class="section" id="tasks">
<details{sch_open}><summary class="{sch_class}"><span class="icon">&#x23F0;</span> Scheduled Tasks {status_badge(sch.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table><tr><th>Label</th><th>PID</th><th>Status</th></tr>
"""
    for t in sch.get("Tasks", [])[:20]:
        html += f"<tr><td>{t.get('Label', '')}</td><td>{t.get('PID', '')}</td><td>{t.get('Status', '')}</td></tr>"
    html += "</table></div></details></div>"

    # BROWSER EXTENSIONS
    be = results.get("BrowserExtensions", {})
    be_open = "" if be.get("Status") in ("PASS", "CLEAN") else " open"
    be_class = heading_class(be.get("Status", "INFO"))
    html += f"""
<div class="section" id="extensions">
<details{be_open}><summary class="{be_class}"><span class="icon">&#x1F527;</span> Browser Extensions {status_badge(be.get("Status", "INFO"))}</summary>
<div class="section-body">
    <h3 style="margin:10px 0 5px;font-size:0.95em;">Browser Extensions</h3>
    <table><tr><th>Browser</th><th>Profile</th><th>Name</th><th>ID</th></tr>
"""
    if be.get("Extensions"):
        for ext in be["Extensions"][:30]:
            html += f"<tr><td>{ext.get('Browser', '')}</td><td>{ext.get('Profile', '')}</td><td>{ext.get('Name', '')}</td><td><code>{ext.get('ID', '')}</code></td></tr>"
        if len(be.get("Extensions", [])) > 30:
            html += f"<tr><td class='detail' colspan='4'>... and {len(be.get('Extensions', [])) - 30} more</td></tr>"
    else:
        html += "<tr><td class='detail' colspan='4'>No supported browser extensions found or browser data inaccessible.</td></tr>"
    if be.get("SafariExtensions"):
        html += f"<tr><td class='detail' colspan='4'>{len(be.get('SafariExtensions', []))} Safari extension item(s) detected.</td></tr>"
    html += "</table></div></details></div>"

    # SERVICE STATUS
    sv = results.get("ServiceStatus", {})
    sv_open = "" if sv.get("Status") in ("PASS", "CLEAN") else " open"
    sv_class = heading_class(sv.get("Status", "INFO"))
    html += f"""
<div class="section" id="services">
<details{sv_open}><summary class="{sv_class}"><span class="icon">&#x2699;</span> Service Status {status_badge(sv.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table><tr><th>Service</th><th>Installed/Detected</th><th>Running</th></tr>
"""
    for s in sv.get("Services", []):
        running = "Yes" if s.get("Running") else '<span class="detail">No</span>'
        installed = "Yes" if s.get("Installed") else '<span class="detail">No</span>'
        html += f"<tr><td>{s.get('Name', '')}</td><td>{installed}</td><td>{running}</td></tr>"
    html += "</table></div></details></div>"

    # NETWORK CONFIG
    net = results.get("NetworkConfig", {})
    net_open = "" if net.get("Status") in ("PASS", "CLEAN") else " open"
    net_class = heading_class(net.get("Status", "INFO"))
    html += f"""
<div class="section" id="network">
<details{net_open}><summary class="{net_class}"><span class="icon">&#x1F310;</span> Network Configuration {status_badge(net.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table>
        {"<tr><td><strong>Wi-Fi Network</strong></td><td>" + net.get('WiFi', 'Not connected') + "</td></tr>" if net.get('WiFi') else ""}
        {"<tr><td><strong>Default Gateway</strong></td><td>" + net.get('DefaultGateway', 'Unknown') + "</td></tr>" if net.get('DefaultGateway') else ""}
        <tr><td><strong>DNS Servers</strong></td><td>{', '.join(net.get('DNS', ['Unknown']))}</td></tr>
    </table>
</div></details></div>
"""

    # MAINTENANCE AGENT HEALTH
    ah = results.get("AgentHealth", {})
    ah_open = "" if ah.get("Status") in ("PASS", "CLEAN") else " open"
    ah_class = heading_class(ah.get("Status", "INFO"))
    html += f"""
<div class="section" id="agent">
<details{ah_open}><summary class="{ah_class}"><span class="icon">&#x1F527;</span> PC Master Class Maintenance Agent {status_badge(ah.get("Status", "INFO"))}</summary>
<div class="section-body">
    <table><tr><th>Check</th><th>Result</th><th>Detail</th></tr>
"""
    for c in ah.get("Checks", []):
        ok = "PASS" if c.get("OK") else "WARNING"
        html += f"<tr><td>{c.get('Name', '')}</td><td>{status_badge(ok)}</td><td>{c.get('Detail', '')}</td></tr>"
    if ah.get("RemoteVersion"):
        html += f"<tr><td>Remote version</td><td colspan='2'>{ah.get('RemoteVersion')}</td></tr>"
    html += "</table></div></details></div>"

    # Footer and closing
    html += f"""
<div class="footer">
    Report generated by PC Master Class macOS Maintenance Script v{script_version}<br>
    &copy; {datetime.now().year} PC Master Class — pc-masterclass.com.au
</div>

</div>
</body>
</html>
"""
    
    return html


# ============================================================================
# EMAIL MODULE
# ============================================================================
def send_email(report_html, to_addr, smtp_config, log_file=None, subject_computer_name=""):
    """Send the HTML report via SMTP."""
    try:
        msg = MIMEMultipart('alternative')
        subject_name = subject_computer_name or platform.node()
        msg['Subject'] = f"Maintenance Report — {subject_name}" if "test" not in sys.argv else f"[TEST] Maintenance Report — {subject_name}"
        msg['From'] = smtp_config.get("email_from", smtp_config["smtp_user"])
        msg['To'] = to_addr

        envelope_recipients = [to_addr]
        if to_addr.lower() == "reports@pcmasterclass.com.au":
            # reports@ is an alias of Paul's own mailbox. Gmail accepts the SMTP
            # send, but self-sent alias mail can land only in All Mail/Sent and
            # not surface in Inbox. Bcc Paul's real mailbox so the report is
            # visible while keeping the report addressed to reports@.
            envelope_recipients.append("paul@pcmasterclass.com.au")

        # Attach HTML
        msg.attach(MIMEText(report_html, 'html', 'utf-8'))

        with smtplib.SMTP(smtp_config["smtp_server"], smtp_config["smtp_port"]) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(smtp_config["smtp_user"], smtp_config["smtp_password"])
            server.sendmail(msg['From'], envelope_recipients, msg.as_string())

        return {"Status": "Sent", "To": to_addr, "EnvelopeRecipients": ", ".join(envelope_recipients), "Server": smtp_config["smtp_server"]}
    except Exception as e:
        return {"Status": "ERROR", "Error": str(e)}


# ============================================================================
# MAIN
# ============================================================================
def main():
    global logger
    
    parser = argparse.ArgumentParser(description="PC Master Class macOS Maintenance Script")
    parser.add_argument("--skip-updates", action="store_true", help="Skip software update check")
    parser.add_argument("--skip-smart", action="store_true", help="Skip SMART disk health check")
    parser.add_argument("--client-name", default="", help="Client name for report header")
    parser.add_argument("--computer-name", default="", help="Optional friendly computer/device name override for report header")
    parser.add_argument("--report-path", default=str(Path.home() / "Library/PCMasterClass/Reports"), help="Report output directory")
    parser.add_argument("--email-to", default="", help="Send report via email")
    parser.add_argument("--smtp-user", default="", help="SMTP username")
    parser.add_argument("--smtp-password", default="", help="SMTP password (App Password)")
    parser.add_argument("--smtp-server", default="smtp.gmail.com", help="SMTP server")
    parser.add_argument("--smtp-port", type=int, default=587, help="SMTP port")
    parser.add_argument("--email-from", default="", help="Sender email address")
    parser.add_argument("--save-credential", action="store_true", help="Store credentials in macOS Keychain")
    parser.add_argument("--skip-update-check", action="store_true", help="Skip auto-update from GitHub")
    args = parser.parse_args()

    # Ensure report directory exists
    report_path = Path(args.report_path)
    report_path.mkdir(parents=True, exist_ok=True)

    # Setup logger
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = report_path / f"{platform.node()}_Maintenance_{timestamp}.log"
    logger = Logger(str(log_file))

    logger.info("=" * 50)
    logger.info(f"PC Master Class macOS Maintenance v{SCRIPT_VERSION}")
    logger.info(f"Computer: {platform.node()}")
    logger.info("=" * 50)

    # Auto-update check
    if not args.skip_update_check:
        check_and_update()

    # Credential management
    smtp_config = None
    if args.save_credential:
        if not args.smtp_user or not args.smtp_password:
            logger.error("--smtp-user and --smtp-password are required when saving credentials.")
            sys.exit(1)
        try:
            keychain_store(
                args.smtp_user, args.smtp_password,
                args.smtp_server, args.smtp_port,
                args.email_from or args.smtp_user
            )
            logger.info("Credentials saved to macOS Keychain successfully.")
            print("Credentials saved. You can now run without --smtp-user/--smtp-password.")
            sys.exit(0)
        except Exception as e:
            logger.error(f"Failed to save credentials: {e}")
            sys.exit(1)

    if args.email_to:
        if args.smtp_user and args.smtp_password:
            smtp_config = {
                "smtp_user": args.smtp_user,
                "smtp_password": args.smtp_password,
                "smtp_server": args.smtp_server,
                "smtp_port": args.smtp_port,
                "email_from": args.email_from or args.smtp_user,
            }
        else:
            # Try loading from Keychain. reports@ is an alias of Paul's mailbox,
            # so prefer the real SMTP login credential stored under paul@.
            loaded = load_smtp_credentials(args.email_to)
            if loaded:
                smtp_config = loaded
                logger.info(f"Loaded SMTP credentials from Keychain for {loaded['smtp_user']}")
            else:
                logger.error("No SMTP credentials found. Run with --save-credential first.")

    # Run all modules
    results = {}
    results["SystemInfo"] = run_system_info()
    results["DiskHealth"] = run_disk_health()
    if not args.skip_updates:
        results["SoftwareUpdates"] = run_software_updates()
    else:
        results["SoftwareUpdates"] = {"Status": "SKIPPED"}
    results["Backup"] = run_backup_check()
    results["Malwarebytes"] = run_malwarebytes_check()
    results["Security"] = run_security_check()
    results["LogErrors"] = run_log_errors()
    results["PendingReboot"] = run_pending_reboot()
    results["Firewall"] = run_firewall_check()
    results["UserAccounts"] = run_user_accounts()
    results["TempFiles"] = run_temp_files()
    results["Startup"] = run_startup_check()
    results["ScheduledTasks"] = run_scheduled_tasks()
    results["BrowserExtensions"] = run_browser_extensions()
    results["ServiceStatus"] = run_service_status()
    results["NetworkConfig"] = run_network_config()
    results["AgentHealth"] = run_agent_health(args.email_to)

    # Generate report
    display_computer_name = build_hardware_display_name(results.get("SystemInfo", {}), args.computer_name)
    logger.info(f"Report display computer: {display_computer_name}")
    report_html = generate_html_report(results, args.client_name, args.computer_name, SCRIPT_VERSION)
    report_file = report_path / f"{platform.node()}_Maintenance_{timestamp}.html"
    report_file.write_text(report_html, encoding='utf-8')
    logger.info(f"Report saved to: {report_file}")

    # Send email
    email_result = {"Status": "Not configured"}
    if args.email_to and smtp_config:
        logger.info("Sending report via email...")
        email_result = send_email(report_html, args.email_to, smtp_config, str(log_file), display_computer_name)
        logger.info(f"Email result: {email_result['Status']}")
        if email_result.get("Status") == "ERROR":
            logger.error(f"Email error detail: {email_result.get('Error', 'Unknown email error')}")

    # Print summary
    warnings = sum(1 for r in results.values() if isinstance(r, dict) and r.get("Status") == "WARNING")
    errors = sum(1 for r in results.values() if isinstance(r, dict) and r.get("Status") == "ERROR")
    logger.info("=" * 50)
    logger.info(f"Maintenance complete. Warnings: {warnings}, Errors: {errors}")
    logger.info(f"Report: {report_file}")
    if email_result.get("Status") == "Sent":
        logger.info(f"Email sent to {args.email_to}")
    logger.info("=" * 50)

    # Return overall status code (non-zero if errors)
    sys.exit(1 if errors > 0 else 0)


if __name__ == "__main__":
    main()
