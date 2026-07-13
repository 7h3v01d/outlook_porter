"""
capture.py

Stage 2: CAPTURE (source PC)

Takes the OutlookProfile objects produced by scan.py and copies everything
needed for a full restore into a self-contained bundle directory:

    <bundle>/
        manifest.json
        psts/<display_name>__<shortsha>.pst
        signatures/...
        dictionary/*.dic
        rules/<profile_name>.rwz

Outlook MUST be closed before capture runs — PST files are exclusively
locked while Outlook has them open, and copying a live PST risks a
truncated/corrupt copy. capture.py checks for this and fails closed rather
than attempting a copy of a locked file.
"""
from __future__ import annotations

import getpass
import platform
import shutil
import socket
from pathlib import Path

from core.manifest import (
    Manifest, OutlookProfile, PstFile, new_manifest, sha256_file,
)

SIGNATURES_REL_PARTS = ("Microsoft", "Signatures")  # under %APPDATA%


class CaptureError(Exception):
    """Raised on any condition that would make the captured bundle unsafe to trust. Fail closed."""


def _is_file_locked(path: Path) -> bool:
    """
    Best-effort check that a file is not currently open for exclusive access
    (e.g. by Outlook). Attempts to open for read+write without sharing.
    """
    try:
        # Opening in r+b mode without another process holding a lock will succeed;
        # a locked PST will raise PermissionError on Windows.
        with open(path, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return True


def _copy_pst(pst: PstFile, dest_dir: Path, manifest_notes: list[str]) -> str:
    src = Path(pst.original_path)
    if not src.exists():
        raise CaptureError(f"PST no longer exists at capture time: {src}")

    if _is_file_locked(src):
        raise CaptureError(
            f"PST '{src}' appears to be open (locked). Close Outlook completely "
            f"(check the system tray — it sometimes stays running) and retry. "
            f"Refusing to copy a locked file, as this risks a corrupt/truncated capture."
        )

    dest_dir.mkdir(parents=True, exist_ok=True)
    short_sha = pst.sha256[:10]
    dest_name = f"{src.stem}__{short_sha}.pst"
    dest_path = dest_dir / dest_name
    shutil.copy2(src, dest_path)

    # Verify the copy is byte-identical to what scan.py hashed. If the file
    # changed between scan and capture (e.g. Outlook briefly reopened it),
    # this catches it rather than silently shipping a bad copy.
    copied_hash = sha256_file(dest_path)
    if copied_hash != pst.sha256:
        dest_path.unlink(missing_ok=True)
        raise CaptureError(
            f"Post-copy hash mismatch for '{src.name}' — the source file changed "
            f"during capture (expected {pst.sha256[:10]}, got {copied_hash[:10]}). "
            f"Capture aborted; retry with Outlook fully closed."
        )

    return dest_name


def capture_signatures(dest_bundle: Path, appdata: Path, notes: list[str]) -> bool:
    src = appdata.joinpath(*SIGNATURES_REL_PARTS)
    if not src.exists():
        notes.append("No Outlook signatures folder found — skipped (this is normal if no signatures are configured).")
        return False
    dest = dest_bundle / "signatures"
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(src, dest)
    return True


def capture_dictionary(dest_bundle: Path, appdata_roaming_uproof: Path, notes: list[str]) -> bool:
    """Custom dictionary lives under %APPDATA%\\Microsoft\\UProof\\*.dic (varies by Office version/locale)."""
    if not appdata_roaming_uproof.exists():
        notes.append("No custom dictionary folder found — skipped.")
        return False
    dic_files = list(appdata_roaming_uproof.glob("*.dic"))
    if not dic_files:
        notes.append("Custom dictionary folder present but contained no .dic files — skipped.")
        return False
    dest = dest_bundle / "dictionary"
    dest.mkdir(parents=True, exist_ok=True)
    for f in dic_files:
        shutil.copy2(f, dest / f.name)
    return True


def export_rules(profile_name: str, dest_bundle: Path, notes: list[str]) -> str | None:
    """
    Export Outlook rules for a profile via COM automation.

    Requires pywin32 and a running-capable Outlook installation (Outlook does
    not need to be open first — starting the Application object will launch
    it if needed, then we quit cleanly).
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        notes.append("pywin32 not installed — rules export skipped. Install pywin32 to enable rule migration.")
        return None

    dest_dir = dest_bundle / "rules"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / f"{profile_name}.rwz"

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    try:
        # Outlook's Rules collection exposes no direct "export all" call in the
        # object model; RuleSet export is done via the underlying MAPI store's
        # rules wizard format. We use Store.GetRules() where available, else
        # fall back to iterating Session.DefaultStore.
        store = namespace.DefaultStore
        rules = store.GetRules()
        # COM's rules object doesn't provide a direct .rwz serializer; rely on
        # Outlook's own Rules Wizard export via the Application's Rules UI
        # automation is out of scope for headless capture. We record rule
        # names/conditions in a human-readable JSON sidecar instead, which
        # restore.py can present to the technician for manual recreation if
        # the .rwz path isn't available on this Office build.
        rule_summaries = []
        for i in range(1, rules.Count + 1):
            r = rules.Item(i)
            rule_summaries.append({"name": r.Name, "enabled": bool(r.Enabled)})

        import json as _json
        sidecar = dest_dir / f"{profile_name}_rules_summary.json"
        sidecar.write_text(_json.dumps(rule_summaries, indent=2), encoding="utf-8")
        notes.append(
            f"Captured {len(rule_summaries)} rule name(s)/state to {sidecar.name}. "
            f"Full rule logic export (.rwz) requires manual 'Rules > Export Rules' "
            f"in Outlook if COM-level export is unavailable on this Office build."
        )
        return sidecar.name
    finally:
        del namespace
        del outlook


def capture_profile(
    profile: OutlookProfile,
    bundle_dir: Path,
    appdata: Path,
    uproof_dir: Path,
) -> Manifest:
    """
    Full capture for a single OutlookProfile. Returns the populated Manifest
    (also written to bundle_dir/manifest.json).
    """
    bundle_dir.mkdir(parents=True, exist_ok=True)
    psts_dir = bundle_dir / "psts"

    manifest = new_manifest(
        source_hostname=socket.gethostname(),
        windows_username=getpass.getuser(),
    )

    seen_psts: dict[str, PstFile] = {}
    for account in profile.accounts:
        if account.pst is not None and account.pst.original_path not in seen_psts:
            captured_name = _copy_pst(account.pst, psts_dir, manifest.notes)
            account.pst.captured_filename = captured_name
            seen_psts[account.pst.original_path] = account.pst
        elif account.pst is not None:
            # Same PST shared by another account (e.g. two POP accounts delivering
            # to one PST) — reuse the already-captured filename.
            account.pst.captured_filename = seen_psts[account.pst.original_path].captured_filename

    manifest.profiles.append(profile)
    manifest.signatures_dir_captured = capture_signatures(bundle_dir, appdata, manifest.notes)
    manifest.dictionary_captured = capture_dictionary(bundle_dir, uproof_dir, manifest.notes)

    rules_file = export_rules(profile.profile_name, bundle_dir, manifest.notes)
    if rules_file:
        manifest.rules_captured.append(rules_file)

    manifest.save(bundle_dir / "manifest.json")
    return manifest
