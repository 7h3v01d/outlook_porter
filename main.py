r"""
Outlook Porter — CLI entry point

Usage:
    python main.py scan
    python main.py capture --bundle "D:\OutlookMigration\bundle"
    python main.py restore --bundle "D:\OutlookMigration\bundle" --pst-dir "C:\Users\<user>\Documents\Outlook Files"
    python main.py verify  --bundle "D:\OutlookMigration\bundle" --pst-dir "C:\Users\<user>\Documents\Outlook Files"

This is a thin CLI over core/. The PyQt6 GUI (gui/main_window.py) is the
intended day-to-day interface for the technician; the CLI exists for
scripting/automation and for testing the pipeline without the GUI.
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from core.manifest import Manifest
from core.scan import scan_all_profiles, ScanError
from core.capture import capture_profile, CaptureError
from core.restore import restore_profile, run_importprf, RestoreError
from core.verify import run_restore_verify, VerifyError


def cmd_scan(args: argparse.Namespace) -> int:
    notes: list[str] = []
    try:
        profiles = scan_all_profiles(notes=notes)
    except ScanError as exc:
        print(f"[SCAN FAILED] {exc}", file=sys.stderr)
        return 1

    for profile in profiles:
        print(f"Profile: {profile.profile_name} (Outlook {profile.outlook_version})")
        for account in profile.accounts:
            pst_info = account.pst.original_path if account.pst else "(no PST)"
            print(f"  - {account.account_name} <{account.email_address}> [{account.account_type}] -> {pst_info}")
        for archive in profile.archive_psts:
            print(f"  - (archive PST) {archive.display_name} -> {archive.original_path}")
    if notes:
        print("\nDiagnostics:")
        for note in notes:
            print(f"  note: {note}")
    return 0


def cmd_capture(args: argparse.Namespace) -> int:
    bundle_dir = Path(args.bundle)
    try:
        scan_notes: list[str] = []
        profiles = scan_all_profiles(notes=scan_notes)
        if len(profiles) > 1:
            print(f"Multiple profiles found ({[p.profile_name for p in profiles]}); "
                  f"capturing '{args.profile or profiles[0].profile_name}'. Use --profile to choose.")
        target = next((p for p in profiles if p.profile_name == args.profile), profiles[0])

        appdata = Path(os.environ.get("APPDATA", ""))
        uproof = appdata / "Microsoft" / "UProof"

        manifest = capture_profile(target, bundle_dir, appdata, uproof)
        # capture_profile builds its own manifest and notes list internally
        # (it doesn't know about scan-time diagnostics) — merge scan_notes
        # in so they end up saved in manifest.json instead of only ever
        # being printed once to the console and lost.
        manifest.notes = scan_notes + manifest.notes
        manifest.save(bundle_dir / "manifest.json")
        print(f"Capture complete -> {bundle_dir}")
        for note in manifest.notes:
            print(f"  note: {note}")
        return 0
    except (ScanError, CaptureError) as exc:
        print(f"[CAPTURE FAILED] {exc}", file=sys.stderr)
        return 1


def cmd_restore(args: argparse.Namespace) -> int:
    bundle_dir = Path(args.bundle)
    pst_dir = Path(args.pst_dir)
    try:
        manifest = Manifest.load(bundle_dir / "manifest.json")
        prf_path = bundle_dir / "restore.prf"
        restore_profile(manifest, bundle_dir, pst_dir, prf_path)
        print(f"PST(s) restored to {pst_dir}")
        print(f"PRF written to {prf_path}")
        if args.run_outlook:
            run_importprf(prf_path, outlook_exe=args.outlook_exe)
            print("Outlook launched with /importprf.")
        else:
            print(f"Run manually: \"{args.outlook_exe}\" /importprf \"{prf_path}\"")
        return 0
    except RestoreError as exc:
        print(f"[RESTORE FAILED] {exc}", file=sys.stderr)
        return 1


def cmd_verify(args: argparse.Namespace) -> int:
    bundle_dir = Path(args.bundle)
    pst_dir = Path(args.pst_dir)
    try:
        manifest = Manifest.load(bundle_dir / "manifest.json")
        restored_paths = {}
        for account in manifest.profiles[0].accounts:
            if account.pst:
                restored_paths[account.pst.original_path] = pst_dir / f"{account.pst.display_name}.pst"

        discrepancies = run_restore_verify(manifest, restored_paths)
        if discrepancies:
            print("VERIFY: discrepancies found:")
            for d in discrepancies:
                print(f"  - {d}")
            return 1
        print("VERIFY: clean pass — restored data matches captured baseline.")
        return 0
    except VerifyError as exc:
        print(f"[VERIFY FAILED] {exc}", file=sys.stderr)
        return 1


def cmd_gui(args: argparse.Namespace) -> int:
    from gui.main_window import main as gui_main
    return gui_main()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="outlook-porter")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("gui", help="Launch the PyQt6 GUI").set_defaults(func=cmd_gui)

    sub.add_parser("scan", help="List Outlook profiles/accounts on this PC").set_defaults(func=cmd_scan)

    p_capture = sub.add_parser("capture", help="Capture PST/signatures/rules into a bundle")
    p_capture.add_argument("--bundle", required=True, help="Output bundle directory")
    p_capture.add_argument("--profile", default=None, help="Profile name to capture (default: first found)")
    p_capture.set_defaults(func=cmd_capture)

    p_restore = sub.add_parser("restore", help="Restore a bundle to this PC")
    p_restore.add_argument("--bundle", required=True, help="Bundle directory (from capture)")
    p_restore.add_argument("--pst-dir", required=True, help="Where to place restored PST files")
    p_restore.add_argument("--run-outlook", action="store_true", help="Actually launch outlook.exe /importprf")
    p_restore.add_argument("--outlook-exe", default="outlook.exe", help="Path to outlook.exe if not on PATH")
    p_restore.set_defaults(func=cmd_restore)

    p_verify = sub.add_parser("verify", help="Restore-verify a completed restore")
    p_verify.add_argument("--bundle", required=True)
    p_verify.add_argument("--pst-dir", required=True)
    p_verify.set_defaults(func=cmd_verify)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
