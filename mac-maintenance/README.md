# PC Master Class — macOS Periodic Maintenance

A comprehensive, automated maintenance and health-check system for Apple Mac computers, matching the architecture and reporting style of the existing Windows PC maintenance program.

## Overview

- **Languages:** Python 3 (built into macOS 11+)
- **Scheduling:** macOS LaunchAgent (native alternative to Task Scheduler)
- **Reporting:** Branded HTML reports with collapsible sections, emailed via Gmail SMTP
- **Updates:** Auto-pulls latest script from GitHub before each run
- **Security:** Credentials stored in macOS Keychain (encrypted)

## Files

| File | Purpose |
|------|---------|
| `pcm_mac_maintenance.py` | Main maintenance script (all checks, HTML report, email) |
| `install.sh` | One-click installer — downloads script, installs LaunchAgent, sets credentials |
| `com.pcmasterclass.maintenance.agent.plist` | LaunchAgent template for quarterly scheduling |
| `docs/apple-security-guide.md` | Detailed Apple security hardening and deployment notes |

## Quick Start (for Paul / Remote Deployment)

1. **Via TeamViewer** on target Mac:
   ```bash
   # Open Terminal and run:
   curl -fsSL https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/install.sh | bash
   ```

2. **Follow prompts:**
   - Grant Full Disk Access (explained in installer)
   - Enter Gmail App Password
   - Enter recipient email (paul@pcmasterclass.com.au)

3. **Done.** The script runs its first test scan immediately, then quarterly forever.

## Manual Run

```bash
python3 ~/Library/PCMasterClass/pcm_mac_maintenance.py \
    --email-to paul@pcmasterclass.com.au \
    --client-name "Client Surname"
```

## Credentials

The installer saves SMTP credentials using the built-in **macOS Keychain**, which is encrypted with the user's login password. Unlike the Windows script (which needed DPAPI + a fallback AES key for SYSTEM tasks), the Mac Keychain works for both interactive runs and LaunchAgent background execution without special handling.

## Architecture vs Windows Version

| Component | Windows (PC) | macOS (Mac) |
|---|---|---|
| Language | PowerShell | Python 3 |
| Scheduler | Task Scheduler | LaunchAgent |
| Identity | Windows DPAPI + AES fallback | macOS Keychain |
| Auto-update | Downloads from GitHub, self-replaces | Same |
| Report style | HTML, collapsible, branded | Same CSS/styling |
| Module count | 21 | 15 (merged/missing equivalents) |
| Notable missing | SFC, DISM, AdwCleaner, Windows Defender | No exact equivalents; replaced by SIP + XProtect checks |

## What the Script Checks

1. **System Info** — Model, CPU, RAM, serial, FileVault, uptime
2. **Disk Health** — Volume usage, SMART (if smartmontools installed)
3. **Software Updates** — Apple softwareupdate + Homebrew outdated
4. **Backup** — iDrive (if installed) + Time Machine
5. **Malwarebytes** — Endpoint Protection status
6. **macOS Security** — Gatekeeper, XProtect, SIP
7. **Log Errors** — unified log errors + crash reports
8. **Pending Reboot** — Uptime check, pending update reboot flags
9. **Firewall** — Application Firewall on/off
10. **User Accounts** — Local accounts with admin status
11. **Temp Files** — Cache directory sizes
12. **Startup Items** — Login items, LaunchAgents
13. **Scheduled Tasks** — Running LaunchAgents/Daemons
14. **Browser Extensions** — Chrome extensions (Safari is private API)
15. **Services** — Key system services (WindowServer, Bluetooth, etc.)
16. **Network** — WiFi name, gateway, DNS

## Differences from Windows Version

- **No SFC/DISM:** macOS handles file integrity differently; there is no online repair mechanism.
- **No AdwCleaner:** Tool is Windows-only. Malwarebytes for Mac provides similar adware cleanup.
- **Windows Defender → XProtect/Gatekeeper:** macOS's built-in malware protection is signature-based (XProtect) and checks Gatekeeper enforcement.
- **Event Log → Unified Logging:** `log show` is verbose and requires filtering. The script looks for error-level entries and crash reports instead.
- **Task Scheduler → LaunchAgent:** macOS LaunchAgents are user-scoped. For true system-level background execution, a LaunchDaemon at `/Library/LaunchDaemons/` (requiring admin) would be needed.
- **iDrive + Time Machine:** Mac clients may use either. The script checks both.

## Security & Apple Gatekeeping

See `docs/apple-security-guide.md` for detailed handling of:
- Full Disk Access (required for browser extension check, user caches)
- Gatekeeper / quarantine flag removal
- Notarization (not needed for Python scripts but needed for bundled apps)
- Code signing (not needed for script-based installation)

## License

Proprietary — Paul Benjamin, PC Master Class
