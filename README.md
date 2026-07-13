# Outlook Porter

A standalone Windows tool for migrating **POP3/IMAP** Outlook accounts
between PCs. Built for client PC-migration engagements where a 1:1 Outlook
setup is expected on the new machine.

## Scope (read this first)

This tool supports **POP3 and IMAP accounts only.** Exchange / Microsoft 365
(cloud) accounts are out of scope by design — those accounts should simply
be re-added on the new PC and left to resync from the server; attempting to
migrate an OST (the local cache file) is unsupported by Microsoft and
produces unreliable results.

### What gets migrated
- PST files (mail, calendar, contacts, tasks — everything in the data file)
- Account configuration (servers, ports, SSL, username) — recreated via a
  Microsoft-supported `.prf` profile script, **not** raw registry cloning
- Signatures
- Custom dictionary
- Rule names/enabled-state (full rule *logic* export depends on Office
  build — see Known Limitations)

### What does NOT get migrated — and why
**Passwords / saved credentials.** Outlook credentials are protected by
Windows DPAPI, cryptographically tied to the Windows user account and
machine that saved them. They cannot be extracted and replayed on a
different PC — this is a Windows security boundary, not a gap in this tool.

The client will be prompted for their email password once, on first
send/receive, after restore. This is the expected and safe behaviour.
**This tool never reads, stores, or transmits account passwords**, in
keeping with KeystoneAI's fail-closed / no-hidden-execution principles —
there is no code path in `core/` that touches Credential Manager.

## Pipeline

```
scan     (source PC, read-only)   -> lists profiles/accounts/PST locations
capture  (source PC)              -> copies PSTs+signatures+dictionary+rules into a bundle dir, writes manifest.json
restore  (target PC)              -> places PSTs, generates restore.prf
verify   (target PC)              -> hash-checks PSTs; if pywin32/Outlook available, compares folder item counts
```

Each stage is independent and auditable — the bundle directory and
`manifest.json` are plain files a technician can inspect before trusting
them. Restore never silently overwrites an existing file at the
destination (fail closed).

## Usage (CLI)

```bash
# On the OLD pc (Outlook fully closed first):
python main.py scan
python main.py capture --bundle "D:\OutlookMigration\bundle"

# Copy the bundle folder to the new PC (USB, network share, etc.)

# On the NEW pc:
python main.py restore --bundle "D:\OutlookMigration\bundle" --pst-dir "C:\Users\<user>\Documents\Outlook Files"
# Review restore.prf, then either pass --run-outlook or run manually:
outlook.exe /importprf "D:\OutlookMigration\bundle\restore.prf"

# After first launch + password entry:
python main.py verify --bundle "D:\OutlookMigration\bundle" --pst-dir "C:\Users\<user>\Documents\Outlook Files"
```

## Usage (GUI)

```bash
python main.py gui
```

**Quick Migrate** (default tab) is the two-button path for normal jobs:
- **On the old PC:** "Back Up This PC's Outlook..." — pick a folder, and it
  scans, auto-picks the profile (or prompts if there's more than one), and
  captures PST + signatures + dictionary + rules into that folder in one go.
- **On the new PC:** "Restore Outlook Here..." — pick the copied backup
  folder and a destination for the PSTs; it restores the files, generates
  the `.prf`, and launches Outlook automatically (auto-detecting
  `outlook.exe` under Program Files). Once the client has signed in and
  mail has synced a bit, click **"I've signed in — Verify now"**.

**Advanced** contains the original four-tab flow (Scan / Capture / Restore /
Verify) for inspecting the manifest, choosing a specific `outlook.exe` path,
re-running a single stage, or any other case where you want manual control
rather than the Quick Migrate defaults.

All work runs on a background thread so the window never locks up during a
hash or PST copy; the log pane at the bottom mirrors what the CLI would
print either way.

## Known limitations

- **Rules**: Outlook's COM object model doesn't expose a direct `.rwz`
  export call. `capture.py` records rule names + enabled state as a JSON
  sidecar so nothing is silently lost, but full rule *conditions/actions*
  currently need manual recreation (Outlook's own Rules Wizard export, if
  available on the source Office build) — flagged clearly in the manifest
  notes, never silently dropped.
- **Multiple profiles**: `restore.py` currently restores one profile per
  run by design (fail closed rather than guessing which profile should be
  "the" profile on the new machine). If a client has multiple profiles,
  run restore once per profile.
- **Outlook must be fully closed during capture.** `capture.py` detects a
  locked PST and aborts rather than risking a truncated copy — check the
  system tray, Outlook sometimes lingers there after the window closes.
- Requires Outlook Desktop (not New Outlook / Outlook for Windows's
  lightweight mode, which uses a different account model entirely — verify
  which one the client has installed before running scan).

## Tests

```bash
pip install pytest --break-system-packages   # or use a venv
pytest tests/ -v
```

28 tests cover manifest serialization, `.prf` generation, restore's
overwrite-refusal and hash-verification logic, and capture's lock-detection
and fail-closed abort paths. Registry (`winreg`) and Outlook COM
(`win32com`) calls are isolated behind small functions and are not
exercised by this suite (Windows-only) — they're the natural next thing to
validate against a real Outlook install / test VM before client use.
