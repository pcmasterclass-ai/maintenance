# Apple Security Guide for macOS Maintenance Deployment

## 1. Gatekeeper & Quarantine

**Problem:** Files downloaded from the internet get a `com.apple.quarantine` extended attribute. When you double-click them, macOS warns "cannot verify the developer" and refuses to run them.

**Solutions (pick one):**

### Option A — Right-click → Open (easiest, user-friendly)
1. Find the downloaded `install.sh` or `pcm_mac_maintenance.py` in Finder
2. **Right-click** (or Control-click) the file
3. Select **Open**
4. Click **Open** again in the security dialog
5. The file is whitelisted for that user forever after

### Option B — Terminal `xattr` (Paul's preference during TeamViewer)
Since you control the Mac via TeamViewer, just strip the quarantine flag:
```bash
xattr -cr ~/Library/PCMasterClass/pcm_mac_maintenance.py
xattr -cr ~/Library/PCMasterClass/install.sh
```

**What this does:** Removes the `com.apple.quarantine` extended attribute from the file and everything inside the directory. Gatekeeper no longer applies.

### Option C — System Settings override
System Settings → Privacy & Security → Security → "Open Anyway" — shown after a failed launch attempt.

## 2. Full Disk Access (FDA)

**Problem:** Starting with macOS 10.14 (Mojave), Apple requires apps/scripts to request explicit permission to access:
- Desktop, Documents, Downloads
- User Library (caches, browser data, Mail)
- External disks
- System files

If the script runs without FDA, you'll get empty results for:
- Browser extensions
- User cache sizes
- Some log data

**Solution:**
1. System Settings → Privacy & Security → Full Disk Access
2. Click the lock to make changes (authenticate)
3. Add **Terminal** (if running via Terminal) or **Python** (if running as binary)
4. Restart Terminal / the script runner

**Note for LaunchAgent:** Since the LaunchAgent runs as the user, and `python3` is located at `/usr/bin/python3` (part of macOS), it generally inherits the user's permissions. However, some directories (e.g. Safari extensions) may still be inaccessible without FDA.

## 3. Notarization & Code Signing

**Do we need it?**

| Approach | Notarization needed? | Code signing needed? |
|---|---|---|
| Raw Python script + `install.sh` | **No** | **No** |
| `.app` bundle (e.g. PyInstaller) | Yes (on macOS 10.15+) | Yes (to avoid right-click) |
| `.pkg` installer | Yes (recommended) | Yes |
| `.dmg` with disk image | No | No (but the .app inside would need both) |

**Our approach:** We use a **Python script + shell installer**. Python scripts are not subject to Gatekeeper enforcement because the Python interpreter itself is already signed by Apple. As long as the script is not in quarantine (stripped via `xattr -cr`), it runs without additional signing.

**If you decide to package as .app later:**
- You'd need a **Developer ID Application** certificate ($99/year Apple Developer Program)
- Submit to Apple for notarization via `xcrun altool`
- This adds friction but removes all security prompts for end users

## 4. System Integrity Protection (SIP)

SIP is always enabled on modern Macs. The script **checks** SIP status (good hygiene) but never needs to disable it.

SIP protects:
- `/System`, `/sbin`, `/usr` (except `/usr/local`)
- Kernel extensions
- NVRAM variables

Our script operates entirely in user space (`~/Library`, `/tmp`, etc.), so SIP is irrelevant.

## 5. TCC (Transparency, Consent, and Control)

Beyond Full Disk Access, macOS has granular TCC permissions:

| Permission | Needed for | Script impact if denied |
|---|---|---|
| **Accessibility** | Not needed | None |
| **Camera** | Not needed | None |
| **Microphone** | Not needed | None |
| **Screen Recording** | Not needed | None |
| **Location** | Not needed | None |
| **Files & Folders** (Documents, Desktop) | Cache scanning | Minor — will skip those directories |

The script will handle permission denial gracefully and report "N/A" or skip the affected checks.

## 6. Staging a Test Machine

Before deploying to live client Macs, test on your own Mac:

```bash
# 1. Create a test directory
mkdir -p ~/Library/PCMasterClass/Scripts

# 2. Copy the script
cp pcm_mac_maintenance.py ~/Library/PCMasterClass/Scripts/

# 3. Strip quarantine (fresh from GitHub will have it)
xattr -cr ~/Library/PCMasterClass/Scripts/pcm_mac_maintenance.py

# 4. Run with test flag (adds [TEST] to email subject)
python3 ~/Library/PCMasterClass/Scripts/pcm_mac_maintenance.py \
    --email-to reports@pcmasterclass.com.au \
    --client-name "TEST Mac" \
    --report-path ~/Library/PCMasterClass/Reports

# 5. Check the HTML report in your browser
open ~/Library/PCMasterClass/Reports/*.html
```

## 7. Deployment Checklist (per client Mac)

During a TeamViewer session:

- [ ] Open Terminal
- [ ] Download and run installer:
  ```bash
  curl -fsSL https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/install.sh | bash
  ```
- [ ] Grant Full Disk Access to Terminal
- [ ] Enter Gmail App Password when prompted
- [ ] Verify test report arrives in Gmail
- [ ] Verify LaunchAgent is loaded:
  ```bash
  launchctl list | grep pcmasterclass
  ```
- [ ] Close Terminal and disconnect

## 8. Troubleshooting Common Security Blocks

| Symptom | Cause | Fix |
|---|---|---|
| "Cannot be opened because the developer cannot be verified" | Quarantine flag | `xattr -cr <file>` |
| "Operation not permitted" when reading Library | Missing FDA | Add Terminal to FDA |
| Keychain password prompt at every run | Wrong Keychain permissions | Run `security unlock-keychain` first |
| Empty browser extension list | TCC blocking ~/Library access | FDA + restart script runner |
| Script runs but email fails | Gmail blocking "less secure app" | Use an **App Password**, not regular password |

## 9. Comparison: LaunchAgent vs LaunchDaemon vs cron

| Method | User scope | Requires root | Survives login? | Best for |
|---|---|---|---|---|
| **LaunchAgent** (`~/Library/LaunchAgents`) | User | No | No (runs while logged in) | ✅ Our use case |
| **LaunchDaemon** (`/Library/LaunchDaemons`) | System | Yes (admin) | Yes | Background servers |
| **cron** (`crontab -e`) | User | No | Yes (system-wide) | ⚠️ Deprecated on macOS |

Apple has deprecated `cron` in favor of `launchd`. LaunchAgents are the native, Apple-approved way to schedule user-level periodic tasks on macOS. They integrate with `log show`, `launchctl`, and energy management.
