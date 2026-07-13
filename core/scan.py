"""
scan.py

Stage 1: SCAN (source PC)

Reads Outlook profile/account configuration from the registry and locates
associated PST files on disk. Produces MailAccount / OutlookProfile objects
(see manifest.py) without copying anything yet — scan is read-only.

Outlook stores POP3/IMAP account settings under:
  HKCU\\Software\\Microsoft\\Office\\<version>\\Outlook\\Profiles\\<ProfileName>\\9375CFF0413111d3B88A00104B2A6676\\<NNNNNNNN>

Each account subkey is a binary-heavy blob, but the fields we need
(email address, servers, ports, username, SSL flags, associated PST path)
are exposed as individually named values within that subkey — this module
reads those named values directly rather than parsing the binary blob,
which is the same approach Microsoft's own migration tooling uses.

This module is Windows-only at runtime (winreg). It is written so the
registry access is isolated behind small functions that can be mocked in
tests on any platform.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Iterable

from core.manifest import MailAccount, OutlookProfile, PstFile, sha256_file

# Office version -> friendly name, newest first. We probe all of these
# because a machine may have leftover profile data from an older install.
KNOWN_OFFICE_VERSIONS = ["16.0", "15.0", "14.0"]

OUTLOOK_PROFILES_KEY_TEMPLATE = r"Software\Microsoft\Office\{version}\Outlook\Profiles"
ACCOUNT_SUBKEY_GUID = "9375CFF0413111d3B88A00104B2A6676"

# Real Outlook profile data (verified against a live registry dump, 2026-07)
# does NOT have a literal "Account Type" value. Mail accounts are distinguished
# by having an "IMAP Server" or "POP3 Server" value; non-mail services in the
# same key (address book, etc.) have neither. We infer type from whichever
# server field is present rather than trusting a value that doesn't exist.
SUPPORTED_ACCOUNT_TYPES = ("POP3", "IMAP")

# Message-store entry ID binary blobs (e.g. "Delivery Store EntryID") embed
# the on-disk file path as a UTF-16LE string alongside a provider DLL name
# ("pstprx.dll"). There's no separate plain-string "Delivery Store PST" value
# for POP3/IMAP accounts on current Outlook — we have to pull the path out
# of this blob instead.
_PATH_IN_ENTRY_ID_RE = re.compile(
    r"[A-Za-z]:\\(?:[^\\/:*?\"<>|\x00]+\\)*[^\\/:*?\"<>|\x00]+\.(?:pst|ost)",
    re.IGNORECASE,
)

# Where Outlook data files (.pst) normally live outside of an account's
# default delivery location — e.g. manually-added archives or extra data
# files. We search these when building archive_psts.
def _default_pst_search_dirs() -> list[Path]:
    dirs = []
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    if local_appdata:
        dirs.append(Path(local_appdata) / "Microsoft" / "Outlook")
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        dirs.append(Path(userprofile) / "Documents" / "Outlook Files")
    return dirs


def _extract_store_path_from_entry_id(entry_id: bytes) -> str | None:
    """Pull an embedded file path (.pst or .ost) out of a MAPI store entry ID blob."""
    try:
        text = entry_id.decode("utf-16-le", errors="ignore")
    except Exception:
        return None
    match = _PATH_IN_ENTRY_ID_RE.search(text)
    return match.group(0) if match else None


class ScanError(Exception):
    """Raised when the registry/filesystem is in a state we refuse to guess about."""


def _require_windows() -> None:
    if sys.platform != "win32":
        raise ScanError(
            "Outlook profile scanning requires the Windows registry (winreg) "
            "and is not available on this platform."
        )


def _open_winreg():
    _require_windows()
    import winreg  # type: ignore
    return winreg


def list_profile_names(version: str) -> list[str]:
    """Return every Outlook profile name defined for a given Office version, or [] if none."""
    winreg = _open_winreg()
    key_path = OUTLOOK_PROFILES_KEY_TEMPLATE.format(version=version)
    names: list[str] = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            i = 0
            while True:
                try:
                    names.append(winreg.EnumKey(key, i))
                    i += 1
                except OSError:
                    break
    except FileNotFoundError:
        return []
    return names


def _read_string_value(winreg, key, name: str, default: str = "") -> str:
    try:
        value, _ = winreg.QueryValueEx(key, name)
        if isinstance(value, bytes):
            # Outlook stores many "string" values as null-terminated UTF-16LE byte blobs
            return value.decode("utf-16-le", errors="ignore").rstrip("\x00")
        return str(value)
    except FileNotFoundError:
        return default


def _read_dword_value(winreg, key, name: str, default: int = 0) -> int:
    try:
        value, _ = winreg.QueryValueEx(key, name)
        return int(value)
    except FileNotFoundError:
        return default


def _read_bool_value(winreg, key, name: str, default: bool = False) -> bool:
    try:
        value, _ = winreg.QueryValueEx(key, name)
        return bool(value)
    except FileNotFoundError:
        return default


def scan_profile(version: str, profile_name: str, notes: list[str] | None = None) -> OutlookProfile:
    """
    Read one Outlook profile's accounts and associated PSTs.

    Raises ScanError rather than returning a partial/guessed profile if the
    profile key is missing or an account subkey is malformed — fail closed.

    `notes`, if provided, is appended with a line for every account subkey
    that was found but skipped (wrong type, etc.) so a caller can surface
    *why* an account didn't show up instead of it just vanishing silently.
    """
    if notes is None:
        notes = []
    winreg = _open_winreg()
    key_path = OUTLOOK_PROFILES_KEY_TEMPLATE.format(version=version) + f"\\{profile_name}\\{ACCOUNT_SUBKEY_GUID}"

    accounts: list[MailAccount] = []
    known_psts: dict[str, PstFile] = {}

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as accounts_key:
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(accounts_key, i)
                except OSError:
                    break
                i += 1

                with winreg.OpenKey(accounts_key, subkey_name) as acct_key:
                    email = _read_string_value(winreg, acct_key, "Email", "")
                    imap_server = _read_string_value(winreg, acct_key, "IMAP Server", "")
                    pop3_server = _read_string_value(winreg, acct_key, "POP3 Server", "")

                    if imap_server:
                        account_type_raw = "IMAP"
                        incoming_server = imap_server
                    elif pop3_server:
                        account_type_raw = "POP3"
                        incoming_server = pop3_server
                    else:
                        # No IMAP/POP3 server field at all — not a mail account we
                        # handle (address book service, Exchange/modern-auth
                        # account, etc.). Noted rather than silently dropped.
                        if email:
                            notes.append(
                                f"Profile '{profile_name}' subkey '{subkey_name}' has Email={email!r} "
                                f"but no IMAP/POP3 Server field — skipped (likely Exchange/modern-auth, "
                                f"which this tool doesn't support)."
                            )
                        else:
                            service_name = _read_string_value(winreg, acct_key, "Service Name", "")
                            notes.append(
                                f"Profile '{profile_name}' subkey '{subkey_name}' is not a mail account "
                                f"(Service Name={service_name!r}) — skipped."
                            )
                        continue

                    display_name = _read_string_value(winreg, acct_key, "Display Name", email)
                    account_name = _read_string_value(winreg, acct_key, "Account Name", display_name)
                    username = _read_string_value(winreg, acct_key, "POP3 User" if account_type_raw == "POP3" else "IMAP User", "")

                    incoming_port = _read_dword_value(
                        winreg, acct_key, "POP3 Port" if account_type_raw == "POP3" else "IMAP Port",
                        995 if account_type_raw == "POP3" else 993,
                    )
                    incoming_ssl = _read_bool_value(winreg, acct_key, "POP3 Use SSL" if account_type_raw == "POP3" else "IMAP Use SSL", True)

                    outgoing_server = _read_string_value(winreg, acct_key, "SMTP Server", "")
                    outgoing_port = _read_dword_value(winreg, acct_key, "SMTP Port", 587)
                    outgoing_ssl = _read_bool_value(winreg, acct_key, "SMTP Use SSL", True)
                    outgoing_auth = _read_bool_value(winreg, acct_key, "SMTP Authenticate", True)

                    leave_on_server = None
                    if account_type_raw == "POP3":
                        leave_on_server = _read_bool_value(winreg, acct_key, "Leave Mail On Server", False)

                    if not email or not incoming_server:
                        raise ScanError(
                            f"Account subkey '{subkey_name}' in profile '{profile_name}' is missing "
                            f"required fields (email/incoming server). Refusing to guess — fail closed."
                        )

                    # Data-file path isn't a plain string value on current Outlook —
                    # it's embedded in the store entry ID binary blob.
                    store_path_str = None
                    try:
                        entry_id_bytes, _ = winreg.QueryValueEx(acct_key, "Delivery Store EntryID")
                        if isinstance(entry_id_bytes, bytes):
                            store_path_str = _extract_store_path_from_entry_id(entry_id_bytes)
                    except FileNotFoundError:
                        pass

                    pst_obj = None
                    if store_path_str:
                        p = Path(store_path_str)
                        ext = p.suffix.lower()
                        if ext == ".ost":
                            # Cached-mode sync file — not the source of truth, and not
                            # needed for a normal migration: re-adding the account on
                            # the destination PC rebuilds it from the mail server.
                            notes.append(
                                f"Account '{account_name}' ({email}) uses a local cache "
                                f"(.ost) at {p} — not captured. This is expected for "
                                f"IMAP accounts and isn't needed to migrate the account; "
                                f"mail is re-synced from the server once the account is "
                                f"recreated on the destination PC."
                            )
                        elif ext == ".pst":
                            pst_obj = known_psts.get(str(p))
                            if pst_obj is None:
                                if not p.exists():
                                    raise ScanError(
                                        f"PST referenced by account '{account_name}' not found on disk: {p}"
                                    )
                                pst_obj = PstFile(
                                    original_path=str(p),
                                    display_name=p.stem,
                                    size_bytes=p.stat().st_size,
                                    sha256=sha256_file(p),
                                    is_default_delivery=True,
                                )
                                known_psts[str(p)] = pst_obj
                        else:
                            notes.append(
                                f"Account '{account_name}': store path {p} has an unrecognized "
                                f"extension {ext!r} — skipped."
                            )

                    accounts.append(
                        MailAccount(
                            account_name=account_name,
                            email_address=email,
                            display_name=display_name,
                            account_type=account_type_raw,
                            incoming_server=incoming_server,
                            incoming_port=incoming_port,
                            incoming_ssl=incoming_ssl,
                            outgoing_server=outgoing_server,
                            outgoing_port=outgoing_port,
                            outgoing_ssl=outgoing_ssl,
                            outgoing_auth=outgoing_auth,
                            username=username,
                            pst=pst_obj,
                            leave_on_server=leave_on_server,
                        )
                    )
    except FileNotFoundError as exc:
        raise ScanError(f"Profile '{profile_name}' not found for Office version {version}") from exc

    archive_psts = find_archive_psts(known_pst_paths=set(known_psts.keys()), notes=notes)

    return OutlookProfile(
        profile_name=profile_name,
        is_default_profile=False,  # set by caller based on DefaultProfile registry value
        outlook_version=version,
        accounts=accounts,
        archive_psts=archive_psts,
    )


def find_archive_psts(
    known_pst_paths: set[str],
    search_dirs: Iterable[Path] | None = None,
    notes: list[str] | None = None,
) -> list[PstFile]:
    """
    Find .pst files on disk that aren't already accounted for as an
    account's default delivery PST — e.g. manually-added archive files,
    or extra data files opened via File > Open > Outlook Data File.

    This was previously unimplemented (scan_profile always returned
    archive_psts=[]), so any non-default PST was silently dropped from
    every capture regardless of what was actually on the machine.
    """
    if notes is None:
        notes = []
    if search_dirs is None:
        search_dirs = _default_pst_search_dirs()

    # Normalize known paths for comparison (case-insensitive on Windows).
    known_normalized = {os.path.normcase(os.path.abspath(p)) for p in known_pst_paths}

    found: dict[str, PstFile] = {}
    for directory in search_dirs:
        if not directory.exists():
            continue
        for pst_path in directory.glob("*.pst"):
            normalized = os.path.normcase(str(pst_path.resolve()))
            if normalized in known_normalized or normalized in found:
                continue
            try:
                found[normalized] = PstFile(
                    original_path=str(pst_path),
                    display_name=pst_path.stem,
                    size_bytes=pst_path.stat().st_size,
                    sha256=sha256_file(pst_path),
                    is_default_delivery=False,
                )
            except OSError as exc:
                notes.append(f"Could not read candidate archive PST '{pst_path}': {exc}")

    if found:
        notes.append(f"Found {len(found)} archive/non-default PST(s) in {[str(d) for d in search_dirs]}.")

    return list(found.values())


def detect_office_version() -> str:
    """Return the first installed Office version (from KNOWN_OFFICE_VERSIONS) that has any profiles."""
    for version in KNOWN_OFFICE_VERSIONS:
        if list_profile_names(version):
            return version
    raise ScanError(
        "No Outlook profiles found for any known Office version "
        f"({', '.join(KNOWN_OFFICE_VERSIONS)}). Is Outlook installed and configured?"
    )


def scan_all_profiles(version: str | None = None, notes: list[str] | None = None) -> list[OutlookProfile]:
    """
    Convenience entry point: scan every profile for the (auto-detected) Office version.

    `notes`, if provided, collects diagnostic lines (skipped accounts, PST
    search results, etc.) from every profile scanned, so a caller can
    print/save them alongside the result instead of them being lost.
    """
    version = version or detect_office_version()
    names = list_profile_names(version)
    if not names:
        raise ScanError(f"No Outlook profiles found for Office version {version}.")
    return [scan_profile(version, name, notes=notes) for name in names]
