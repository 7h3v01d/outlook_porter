"""
verify.py

Stage 4: RESTORE-VERIFY (target PC)

Two levels of verification:

1. File-level (always available, no Outlook needed): re-hash the restored
   PST and confirm it matches the sha256 captured at scan time. Proves the
   bytes made it across intact.

2. Content-level (requires Outlook + pywin32): open the PST via COM and
   walk folders, comparing item counts against VerificationCounts captured
   pre-migration. Proves Outlook itself can read the store and the data is
   where it should be, not just that the file bytes match.

Any mismatch raises rather than warns — restore-verify is the final gate
before telling the client "you're good to go" (fail closed).
"""
from __future__ import annotations

from pathlib import Path

from core.manifest import Manifest, PstFile, VerificationCounts, sha256_file


class VerifyError(Exception):
    pass


def verify_pst_hash(pst: PstFile, restored_path: Path) -> None:
    if not restored_path.exists():
        raise VerifyError(f"Restored PST not found: {restored_path}")
    actual = sha256_file(restored_path)
    if actual != pst.sha256:
        raise VerifyError(
            f"Hash mismatch for '{pst.display_name}': expected {pst.sha256[:12]}, "
            f"got {actual[:12]}. The restored file does not match the captured original."
        )


def collect_folder_counts(pst_path: Path, display_name: str) -> VerificationCounts:
    """
    Opens a PST via Outlook COM and walks its top-level folders, recording
    item counts. Used both at capture time (baseline) and restore-verify
    time (comparison).
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise VerifyError(
            "pywin32 not installed — cannot perform content-level verification. "
            "Install pywin32, or rely on hash-level verification only."
        ) from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    try:
        namespace.AddStore(str(pst_path))
        store = None
        for i in range(1, namespace.Stores.Count + 1):
            s = namespace.Stores.Item(i)
            if Path(s.FilePath).resolve() == pst_path.resolve():
                store = s
                break
        if store is None:
            raise VerifyError(f"Could not locate added store for {pst_path} in Outlook's session.")

        root = store.GetRootFolder()
        counts: dict[str, int] = {}

        def walk(folder, prefix: str = ""):
            path = f"{prefix}/{folder.Name}" if prefix else folder.Name
            counts[path] = folder.Items.Count
            for sub in folder.Folders:
                walk(sub, path)

        for top_folder in root.Folders:
            walk(top_folder)

        namespace.RemoveStore(root)
        return VerificationCounts(pst_display_name=display_name, folder_counts=counts)
    finally:
        del namespace
        del outlook


def verify_counts_match(baseline: VerificationCounts, current: VerificationCounts) -> list[str]:
    """
    Returns a list of human-readable discrepancies (empty list = perfect match).
    Does not raise itself — caller decides how strict to be (e.g. a couple of
    new emails arriving during migration on an IMAP-mirrored folder is
    expected and not a failure; POP3 delivery-only PSTs should match exactly).
    """
    discrepancies = []
    all_folders = set(baseline.folder_counts) | set(current.folder_counts)
    for folder in sorted(all_folders):
        b = baseline.folder_counts.get(folder)
        c = current.folder_counts.get(folder)
        if b is None:
            discrepancies.append(f"New folder appeared after restore: '{folder}' ({c} items)")
        elif c is None:
            discrepancies.append(f"Folder missing after restore: '{folder}' (expected {b} items)")
        elif b != c:
            discrepancies.append(f"Item count mismatch in '{folder}': expected {b}, found {c}")
    return discrepancies


def run_restore_verify(manifest: Manifest, restored_pst_paths: dict[str, Path]) -> list[str]:
    """
    Full restore-verify pass across every PST in the manifest's single profile.
    Returns a combined list of discrepancy strings; empty means clean pass.
    Always runs hash verification; content verification is attempted and
    skipped gracefully (with a note) if pywin32/Outlook isn't available.
    """
    all_discrepancies: list[str] = []
    profile = manifest.profiles[0]
    seen: set[str] = set()

    for account in profile.accounts:
        pst = account.pst
        if pst is None or pst.original_path in seen:
            continue
        seen.add(pst.original_path)

        restored_path = restored_pst_paths.get(pst.original_path)
        if restored_path is None:
            all_discrepancies.append(f"No restored path provided for PST '{pst.display_name}' — cannot verify.")
            continue

        verify_pst_hash(pst, restored_path)  # raises on mismatch — hard stop

        baseline = next(
            (v for v in manifest.verification if v.pst_display_name == pst.display_name), None
        )
        if baseline is None:
            all_discrepancies.append(
                f"No baseline folder counts captured for '{pst.display_name}' — skipping content-level check."
            )
            continue

        try:
            current = collect_folder_counts(restored_path, pst.display_name)
        except VerifyError as exc:
            all_discrepancies.append(f"Content-level check skipped for '{pst.display_name}': {exc}")
            continue

        all_discrepancies.extend(verify_counts_match(baseline, current))

    return all_discrepancies
