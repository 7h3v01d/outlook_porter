"""
manifest.py

Defines the on-disk manifest format that ties the four pipeline stages
together (scan -> capture -> restore -> restore-verify).

Design principles (KeystoneAI standard):
- fail-closed: missing/ambiguous data raises, never silently guessed
- auditable: manifest is plain JSON, human-readable, versioned
- no hidden execution: manifest never contains passwords or executable content
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

MANIFEST_SCHEMA_VERSION = 2

# Outlook's "SMTP Secure Connection" / "IMAP Secure Connection" registry DWORD.
# This is a DIFFERENT axis to the "Use SSL" flag and is the value that actually
# decides implicit-SSL vs STARTTLS. Verified against a live registry dump
# (2026-07): accounts on port 465 carry 0x1, accounts on port 587 carry 0x2.
SECURE_NONE = 0        # plaintext
SECURE_SSL = 1         # implicit SSL/TLS (typically port 465)
SECURE_STARTTLS = 2    # STARTTLS upgrade on a plaintext port (typically 587)
SECURE_AUTO = 3        # let Outlook decide

SECURE_CONNECTION_NAMES = {
    SECURE_NONE: "None",
    SECURE_SSL: "SSL/TLS (implicit)",
    SECURE_STARTTLS: "STARTTLS",
    SECURE_AUTO: "Auto",
}


class ManifestError(Exception):
    """Raised for any manifest read/write/validation failure. Fail closed."""


@dataclass
class PstFile:
    """A single PST file referenced by an account or as a standalone data file."""
    original_path: str          # absolute path on source PC at capture time
    display_name: str           # name shown in Outlook's folder pane
    size_bytes: int
    sha256: str
    is_default_delivery: bool = False   # was this the account's default delivery PST?
    captured_filename: Optional[str] = None  # filename inside the capture bundle


@dataclass
class MailAccount:
    """A single POP3 or IMAP account extracted from the Outlook profile."""
    account_name: str           # friendly name shown in account list
    email_address: str
    display_name: str
    account_type: str           # "POP3" or "IMAP"
    incoming_server: str
    incoming_port: int
    incoming_ssl: bool
    outgoing_server: str
    outgoing_port: int
    outgoing_ssl: bool
    outgoing_auth: bool         # "my outgoing server requires authentication"
    username: str
    pst: Optional[PstFile] = None       # POP3 delivers into a PST
    leave_on_server: Optional[bool] = None  # POP3-specific

    # --- Transport security (schema v2) ---
    # `outgoing_ssl` alone is NOT enough to reproduce an account: Outlook
    # stores an *additional* "SMTP Secure Connection" DWORD that distinguishes
    # implicit SSL (465) from STARTTLS (587). Dropping it and emitting a bare
    # "use SSL = yes" silently breaks sending on STARTTLS accounts, which is
    # what the previous schema did. Carried verbatim, never inferred.
    outgoing_secure_connection: Optional[int] = None   # see SECURE_* constants
    incoming_secure_connection: Optional[int] = None

    # --- Authentication model (schema v2) ---
    # Modern-auth (OAuth2) accounts — e.g. Gmail — cannot be recreated from a
    # .prf, because there is no password to supply and Google no longer accepts
    # basic auth. Emitting them as ordinary basic-auth IMAP accounts produces
    # an account that looks configured but can never sign in. Flagged here so
    # restore can exclude them and surface them as an explicit manual step.
    uses_oauth: bool = False
    auth_identity_uid: Optional[str] = None


@dataclass
class OutlookProfile:
    profile_name: str
    is_default_profile: bool
    outlook_version: str        # e.g. "16.0" for Outlook 2016/365 desktop
    accounts: list[MailAccount] = field(default_factory=list)
    archive_psts: list[PstFile] = field(default_factory=list)  # non-account PSTs (archives, PSTs opened manually)


@dataclass
class VerificationCounts:
    """Per-PST folder item counts captured at scan/capture time, checked again at restore-verify."""
    pst_display_name: str
    folder_counts: dict[str, int]   # folder path -> item count, e.g. {"Inbox": 1532, "Sent Items": 402}


@dataclass
class Manifest:
    schema_version: int
    created_utc: str
    source_hostname: str
    windows_username: str
    profiles: list[OutlookProfile] = field(default_factory=list)
    signatures_dir_captured: bool = False
    dictionary_captured: bool = False
    rules_captured: list[str] = field(default_factory=list)  # captured rule export filenames, per profile
    verification: list[VerificationCounts] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)  # non-fatal warnings collected during capture

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2, ensure_ascii=False)

    def save(self, path: Path) -> None:
        path.write_text(self.to_json(), encoding="utf-8")

    @staticmethod
    def load(path: Path) -> "Manifest":
        if not path.exists():
            raise ManifestError(f"Manifest not found: {path}")
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise ManifestError(f"Manifest is not valid JSON: {exc}") from exc

        found_version = data.get("schema_version")
        if found_version != MANIFEST_SCHEMA_VERSION:
            hint = ""
            if found_version == 1:
                hint = (
                    " Schema v1 bundles did not capture the SMTP Secure Connection value "
                    "or the OAuth2 auth flag, so a restore from one would misconfigure "
                    "STARTTLS accounts and produce an unusable Gmail/modern-auth account. "
                    "The missing data cannot be reconstructed from the bundle — re-run "
                    "capture on the source PC to produce a v2 bundle."
                )
            raise ManifestError(
                f"Unsupported manifest schema version {found_version!r}; "
                f"expected {MANIFEST_SCHEMA_VERSION}. Refusing to proceed (fail closed).{hint}"
            )

        profiles = []
        for p in data.get("profiles", []):
            accounts = []
            for a in p.get("accounts", []):
                pst = PstFile(**a["pst"]) if a.get("pst") else None
                a = {**a, "pst": pst}
                accounts.append(MailAccount(**a))
            archive_psts = [PstFile(**ap) for ap in p.get("archive_psts", [])]
            p = {**p, "accounts": accounts, "archive_psts": archive_psts}
            profiles.append(OutlookProfile(**p))

        verification = [VerificationCounts(**v) for v in data.get("verification", [])]

        data = {
            **data,
            "profiles": profiles,
            "verification": verification,
        }
        return Manifest(**data)


def new_manifest(source_hostname: str, windows_username: str) -> Manifest:
    return Manifest(
        schema_version=MANIFEST_SCHEMA_VERSION,
        created_utc=datetime.now(timezone.utc).isoformat(),
        source_hostname=source_hostname,
        windows_username=windows_username,
    )


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """Stream-hash a file. Used to verify PST integrity survives copy/restore untouched."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()
