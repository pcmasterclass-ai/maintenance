# PC Master Class macOS Maintenance — Initial Rollout Plan

Purpose: move from local development to controlled production deployment using TeamViewer, with clear stop/go gates before widening the rollout.

## Rollout phases

### Phase 0 — Publish readiness check

Goal: confirm the installer will pull the exact script we just tested.

Before any TeamViewer installation:

- [ ] Validate local files:
  ```bash
  cd ~/.hermes/pcm-maintenance/mac-maintenance
  python3 -m py_compile pcm_mac_maintenance.py
  bash -n install.sh
  plutil -lint com.pcmasterclass.maintenance.agent.plist
  ```
- [ ] Confirm all raw GitHub URLs use `/mac-maintenance/`:
  - `mac-maintenance/pcm_mac_maintenance.py`
  - `mac-maintenance/install.sh`
  - `mac-maintenance/com.pcmasterclass.maintenance.agent.plist`
- [ ] Push/sync the tested local files to `pcmasterclass-ai/maintenance`.
- [ ] Compare local files with GitHub raw downloads before saying the installer is ready.
- [ ] Keep the timestamped local backup from the pre-rollout edit.

Stop/go gate: no deployment until local and GitHub raw files match.

### Phase 1 — Paul's own Mac via TeamViewer-style install

Goal: exercise the exact installer flow that will be used on client Macs, but on Paul's own equipment first.

Target: one Paul-owned Mac reachable by TeamViewer.

Install command:

```bash
curl -fsSL https://raw.githubusercontent.com/pcmasterclass-ai/maintenance/main/mac-maintenance/install.sh | bash
```

During installation:

- [ ] Confirm Terminal opens normally through TeamViewer.
- [ ] Grant Full Disk Access to Terminal when prompted.
- [ ] Save SMTP credential into Keychain using Paul’s Gmail App Password, with `reports@pcmasterclass.com.au` as the sender alias and recipient.
- [ ] Confirm first test scan completes.
- [ ] Confirm report email arrives at `reports@pcmasterclass.com.au`.
- [ ] Confirm LaunchAgent is loaded:
  ```bash
  launchctl list | grep pcmasterclass
  ```
- [ ] Run a manual follow-up scan:
  ```bash
  python3 ~/Library/PCMasterClass/pcm_mac_maintenance.py \
    --skip-update-check \
    --email-to reports@pcmasterclass.com.au \
    --client-name "Paul Test Mac"
  ```
- [ ] Review the HTML report for false positives/noise.

Stop/go gate: no client deployment until the email report is readable, the LaunchAgent is loaded, and no installer friction remains unexplained.

### Phase 2 — One friendly client Mac

Goal: test the process on one real client environment with minimal risk.

Recommended first-client criteria:

- Cooperative/friendly client.
- TeamViewer access already reliable.
- Mac is reasonably current, ideally macOS 12+.
- Preferably has at least one of Malwarebytes Endpoint Protection, iDrive, or Time Machine configured so the important Mac-specific checks are exercised.

Deployment steps:

- [ ] Confirm client consent/context for installing PC Master Class periodic maintenance reporting.
- [ ] Run installer via TeamViewer.
- [ ] Grant Full Disk Access if available.
- [ ] Enter SMTP details and recipient.
- [ ] Confirm test report email arrives.
- [ ] Check the report for:
  - Malwarebytes detection accuracy.
  - iDrive or Time Machine detection accuracy.
  - Browser extension results.
  - Service status usefulness.
  - Any scary false positives.
- [ ] Record outcome in rollout tracker.

Stop/go gate: no five-client rollout until the first real client report is acceptable and any false positives are fixed or documented.

### Phase 3 — Five-client pilot

Goal: cover varied Mac setups before wider rollout.

Choose a mix:

- [ ] One Malwarebytes Endpoint Protection Mac.
- [ ] One iDrive Mac.
- [ ] One Time Machine-only Mac.
- [ ] One Mac with Chrome/Edge/Firefox usage.
- [ ] One older macOS or more complex client setup, if available.

For each pilot client:

- [ ] Install through TeamViewer.
- [ ] Confirm report email.
- [ ] Confirm LaunchAgent loaded.
- [ ] Record result in rollout tracker.
- [ ] Note false positives, missing detection, or deployment friction.

Stop/go gate: continue only if at least 4/5 installs are clean and no high-impact script bug is found.

### Phase 4 — Open rollout

Goal: add all known Mac clients to the quarterly maintenance program.

Before opening:

- [ ] Finalise script fixes from the pilot.
- [ ] Push/sync final version to GitHub.
- [ ] Update README and Apple security guide.
- [ ] Update Google Sheet rollout tracker with macOS candidates/statuses.
- [ ] Prepare a short client-facing wording if Paul wants to mention what is being installed.

## Rollback / uninstall on a client Mac

If the install causes trouble, remove the scheduled agent and local files:

```bash
launchctl unload ~/Library/LaunchAgents/com.pcmasterclass.maintenance.agent.plist 2>/dev/null || true
rm -f ~/Library/LaunchAgents/com.pcmasterclass.maintenance.agent.plist
rm -rf ~/Library/PCMasterClass
```

Optional: remove stored SMTP Keychain entries if needed:

```bash
security delete-generic-password -s "PCMasterClass Maintenance SMTP" -a reports@pcmasterclass.com.au 2>/dev/null || true
security delete-generic-password -s "PCMasterClass Maintenance SMTP Config" -a reports@pcmasterclass.com.au 2>/dev/null || true
```

## Data to record per Mac

- Client name.
- Mac name/hostname.
- macOS version.
- Install date.
- Report received: yes/no.
- LaunchAgent loaded: yes/no.
- Backup detected: iDrive / Time Machine / none.
- Malwarebytes detected: yes/no.
- Issues/notes.
- Next action.
