r"""
restore.py

Stage 3: RESTORE (target PC)

Deliberately does NOT attempt to clone registry profile keys directly —
that approach is fragile across Office builds/versions and can produce a
profile Outlook refuses to open. Instead this uses Microsoft's own
supported provisioning mechanism: a .prf ("profile") script, consumed via

    outlook.exe /importprf "C:\\path\\to\\profile.prf"

PRF FORMAT (this is a real, documented format — do not improvise it)
--------------------------------------------------------------------
A .prf is a 7-section INI file. The previous version of this module invented
its own key names ("Account1IMAPServer=...") which do not exist in the format;
Outlook would not have applied them. The real layout is:

    Section 1  [General]                 profile defaults
    Section 2  [Service List]            MAPI services (PSTs, address book)
    Section 3  [Internet Account List]   AccountN=I_Mail | IMAP_I_Mail
    Section 4  [ServiceN]                values for each MAPI service
    Section 5  [AccountN]                values for each internet account
    Section 6  [<service name>]          property->MAPI mappings   (DO NOT MODIFY)
    Section 7  [I_Mail] / [IMAP_I_Mail]  property->MAPI mappings   (DO NOT MODIFY)

Sections 6 and 7 are fixed boilerplate that Outlook needs in order to know
which MAPI property each Section 4/5 key maps to. They are reproduced verbatim
below.

Transport security: encryption is controlled by SMTPSecureConnection
(PT_LONG, 0x020A) — 1 = implicit SSL (465), 2 = STARTTLS (587). A bare
"use SSL" flag cannot express that difference, which is why the manifest
carries the value through from the registry rather than deriving it.

Passwords are deliberately NOT included — see capture.py docstring / README.
The user enters the password once on first send/receive.

OAuth2 accounts (e.g. Gmail) are EXCLUDED from the .prf and reported as manual
steps: there is no password to supply, the provider rejects basic auth, and
emitting them anyway produces an account that looks configured but can never
sign in.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from core.manifest import (
    Manifest, MailAccount, OutlookProfile, SECURE_NONE, SECURE_SSL,
)

# --- Section 7: internet account property mappings. Verbatim. DO NOT MODIFY. ---
PRF_SECTION_7 = """
; ************************************************************************
; Section 7 - Mapping for internet account properties. DO NOT MODIFY.
; ************************************************************************
[I_Mail]
AccountType=POP3
AccountName=PT_UNICODE,0x0002
DisplayName=PT_UNICODE,0x000B
EmailAddress=PT_UNICODE,0x000C
POP3Server=PT_UNICODE,0x0100
POP3UserName=PT_UNICODE,0x0101
POP3UseSPA=PT_LONG,0x0108
Organization=PT_UNICODE,0x0107
ReplyEmailAddress=PT_UNICODE,0x0103
POP3Port=PT_LONG,0x0104
POP3UseSSL=PT_LONG,0x0105
SMTPServer=PT_UNICODE,0x0200
SMTPUseAuth=PT_LONG,0x0203
SMTPAuthMethod=PT_LONG,0x0208
SMTPUserName=PT_UNICODE,0x0204
SMTPUseSPA=PT_LONG,0x0207
ConnectionType=PT_LONG,0x000F
ConnectionOID=PT_UNICODE,0x0010
SMTPPort=PT_LONG,0x0201
SMTPSecureConnection=PT_LONG,0x020A
ServerTimeOut=PT_LONG,0x0209
LeaveOnServer=PT_LONG,0x1000

[IMAP_I_Mail]
AccountType=IMAP
AccountName=PT_UNICODE,0x0002
DisplayName=PT_UNICODE,0x000B
EmailAddress=PT_UNICODE,0x000C
IMAPServer=PT_UNICODE,0x0100
IMAPUserName=PT_UNICODE,0x0101
IMAPUseSPA=PT_LONG,0x0108
Organization=PT_UNICODE,0x0107
ReplyEmailAddress=PT_UNICODE,0x0103
IMAPPort=PT_LONG,0x0104
IMAPUseSSL=PT_LONG,0x0105
SMTPServer=PT_UNICODE,0x0200
SMTPUseAuth=PT_LONG,0x0203
SMTPAuthMethod=PT_LONG,0x0208
SMTPUserName=PT_UNICODE,0x0204
SMTPUseSPA=PT_LONG,0x0207
ConnectionType=PT_LONG,0x000F
ConnectionOID=PT_UNICODE,0x0010
SMTPPort=PT_LONG,0x0201
SMTPSecureConnection=PT_LONG,0x020A
ServerTimeOut=PT_LONG,0x0209
CheckNewImap=PT_LONG,0x1100
RootFolder=PT_UNICODE,0x1101
"""

# --- Section 6: MAPI service property mappings. Verbatim. DO NOT MODIFY. ---
PRF_SECTION_6 = """
; ************************************************************************
; Section 6 - Mapping for profile properties. DO NOT MODIFY.
; ************************************************************************
[Unicode Personal Folders]
ServiceName=MSUPST MS
Name=PT_UNICODE,0x3001
PathAndFilenameToPersonalFolders=PT_STRING8,0x6700
RememberPassword=PT_BOOLEAN,0x6701
EncryptionType=PT_LONG,0x6702
Password=PT_STRING8,0x6703

[Outlook Address Book]
ServiceName=CONTAB
"""


class RestoreError(Exception):
    pass


def _account_mapping_name(account: MailAccount) -> str:
    """Which Section 7 mapping block this account is declared against."""
    return "IMAP_I_Mail" if account.account_type.upper() == "IMAP" else "I_Mail"


def _restore_pst(pst, bundle_dir: Path, target_pst_dir: Path, notes: list[str]) -> Path:
    if not pst.captured_filename:
        raise RestoreError(f"PST '{pst.display_name}' has no captured_filename — capture stage may be incomplete.")
    src = bundle_dir / "psts" / pst.captured_filename
    if not src.exists():
        raise RestoreError(f"Expected captured PST not found in bundle: {src}")

    target_pst_dir.mkdir(parents=True, exist_ok=True)
    dest = target_pst_dir / f"{pst.display_name}.pst"
    if dest.exists():
        raise RestoreError(
            f"Refusing to overwrite existing file at restore target: {dest}. "
            f"Move/rename it first if this is intentional (fail closed — no silent overwrite)."
        )
    shutil.copy2(src, dest)
    return dest


def partition_accounts(profile: OutlookProfile) -> tuple[list[MailAccount], list[MailAccount]]:
    """
    Split accounts into (provisionable via .prf, manual-only).

    OAuth2/modern-auth accounts go in the manual bucket: a .prf can only express
    a basic-auth account, so emitting one would create an account that appears
    configured but can never authenticate.
    """
    provisionable = [a for a in profile.accounts if not a.uses_oauth]
    manual = [a for a in profile.accounts if a.uses_oauth]
    return provisionable, manual


def build_prf(
    profile: OutlookProfile,
    restored_pst_paths: dict[str, Path],
    profile_name: str | None = None,
) -> str:
    """
    Build .prf content for one profile's POP3/IMAP accounts.

    restored_pst_paths maps original_path -> new absolute Path on the target PC.
    Unlike the previous implementation, this argument is actually used: each
    restored PST becomes a "Unicode Personal Folders" service in Sections 2/4,
    and the default-delivery PST is bound as DefaultStore.
    """
    provisionable, manual = partition_accounts(profile)
    if not provisionable:
        raise RestoreError(
            f"Profile '{profile.profile_name}' has no accounts that can be provisioned via .prf "
            f"({len(manual)} OAuth2/modern-auth account(s) must be added manually). "
            f"Nothing to write."
        )

    name = profile_name or profile.profile_name

    # ---- Section 2/4 inputs: one PST service per restored data file ----
    pst_services: list[tuple[str, str, Path]] = []  # (service_id, display_name, path)
    default_store_service: str | None = None
    seen: set[str] = set()
    for account in profile.accounts:
        pst = account.pst
        if pst is None or pst.original_path in seen:
            continue
        seen.add(pst.original_path)
        restored = restored_pst_paths.get(pst.original_path)
        if restored is None:
            continue
        service_id = f"Service{len(pst_services) + 1}"
        pst_services.append((service_id, pst.display_name, restored))
        if pst.is_default_delivery and default_store_service is None:
            default_store_service = service_id

    lines: list[str] = []
    lines.append("; Generated by Outlook Porter — review before importing.")
    lines.append("; Contains NO passwords by design; the user signs in once after restore.")
    lines.append("")
    lines.append("; ****************************************************")
    lines.append("; Section 1 - Profile defaults")
    lines.append("; ****************************************************")
    lines.append("[General]")
    lines.append("Custom=1")
    lines.append(f"ProfileName={name}")
    lines.append("DefaultProfile=Yes")
    # Fail closed: never clobber a profile that already exists on the target.
    lines.append("OverwriteProfile=No")
    lines.append("ModifyDefaultProfileIfPresent=false")
    lines.append("BackupProfile=False")
    if default_store_service:
        lines.append(f"DefaultStore={default_store_service}")
    lines.append("")

    lines.append("; ****************************************************")
    lines.append("; Section 2 - Services in profile")
    lines.append("; ****************************************************")
    lines.append("[Service List]")
    for service_id, _display_name, _path in pst_services:
        lines.append(f"{service_id}=Unicode Personal Folders")
    lines.append(f"Service{len(pst_services) + 1}=Outlook Address Book")
    lines.append("")

    lines.append("; ****************************************************")
    lines.append("; Section 3 - List of internet accounts")
    lines.append("; ****************************************************")
    lines.append("[Internet Account List]")
    for idx, account in enumerate(provisionable, start=1):
        lines.append(f"Account{idx}={_account_mapping_name(account)}")
    lines.append("")

    lines.append("; ****************************************************")
    lines.append("; Section 4 - Values for each service")
    lines.append("; ****************************************************")
    for service_id, display_name, path in pst_services:
        lines.append(f"[{service_id}]")
        lines.append(f"Name={display_name}")
        lines.append(f"PathAndFilenameToPersonalFolders={path}")
        lines.append("OverwriteExistingService=No")
        lines.append("UniqueService=No")
        lines.append("")
    lines.append(f"[Service{len(pst_services) + 1}]")
    lines.append("OverwriteExistingService=No")
    lines.append("UniqueService=Yes")
    lines.append("")

    lines.append("; ****************************************************")
    lines.append("; Section 5 - Values for each internet account")
    lines.append("; ****************************************************")
    for idx, account in enumerate(provisionable, start=1):
        is_imap = account.account_type.upper() == "IMAP"
        prefix = "IMAP" if is_imap else "POP3"
        secure = (
            account.outgoing_secure_connection
            if account.outgoing_secure_connection is not None
            else (SECURE_SSL if account.outgoing_ssl else SECURE_NONE)
        )
        lines.append(f"[Account{idx}]")
        lines.append(f"AccountName={account.account_name}")
        lines.append(f"DisplayName={account.display_name}")
        lines.append(f"EmailAddress={account.email_address}")
        # Only the keys for THIS account type — the old template emitted both
        # POP3* and IMAP* blocks for every account, pointing POP3Server at an
        # IMAP host.
        lines.append(f"{prefix}Server={account.incoming_server}")
        lines.append(f"{prefix}UserName={account.username}")
        lines.append(f"{prefix}Port={account.incoming_port}")
        lines.append(f"{prefix}UseSSL={1 if account.incoming_ssl else 0}")
        lines.append(f"SMTPServer={account.outgoing_server}")
        lines.append(f"SMTPPort={account.outgoing_port}")
        lines.append(f"SMTPUseAuth={1 if account.outgoing_auth else 0}")
        # The line that fixes STARTTLS accounts (587 -> 2, not "SSL=Yes").
        lines.append(f"SMTPSecureConnection={secure}")
        lines.append("ConnectionType=1")
        if not is_imap and account.leave_on_server is not None:
            lines.append(f"LeaveOnServer={1 if account.leave_on_server else 0}")
        lines.append("")

    lines.append(PRF_SECTION_6.strip())
    lines.append("")
    lines.append(PRF_SECTION_7.strip())
    lines.append("")

    return "\n".join(lines)


def write_prf(content: str, path: Path) -> Path:
    """
    Write .prf as ANSI/CRLF — the format Outlook expects.

    A UTF-8 (especially BOM'd) .prf can be silently misparsed. If an account
    contains characters that can't be represented, fail closed rather than
    writing a mangled file with replacement characters in a server name.
    """
    try:
        data = content.encode("cp1252")
    except UnicodeEncodeError as exc:
        raise RestoreError(
            f"PRF contains characters that cannot be written in ANSI (cp1252): {exc}. "
            f"Outlook's .prf parser is not reliably UTF-8 aware — refusing to write a "
            f"file that may be misread. Check account/display names for non-Latin-1 characters."
        ) from exc
    with open(path, "wb") as f:
        f.write(data.replace(b"\n", b"\r\n"))
    return path


def restore_profile(
    manifest: Manifest,
    bundle_dir: Path,
    target_pst_dir: Path,
    prf_output_path: Path,
) -> Path:
    """
    Restores PST files to target_pst_dir and writes a .prf script to
    prf_output_path. Does NOT invoke outlook.exe itself — that's left to
    the GUI/CLI caller so the technician can review the .prf first
    (auditability: nothing runs unseen).
    """
    if len(manifest.profiles) != 1:
        raise RestoreError(
            f"Expected exactly one profile in manifest, found {len(manifest.profiles)}. "
            f"Multi-profile restore is not yet supported — restore profiles one at a time."
        )
    profile = manifest.profiles[0]

    restored_paths: dict[str, Path] = {}
    for account in profile.accounts:
        if account.pst and account.pst.original_path not in restored_paths:
            dest = _restore_pst(account.pst, bundle_dir, target_pst_dir, manifest.notes)
            restored_paths[account.pst.original_path] = dest

    prf_content = build_prf(profile, restored_paths)
    write_prf(prf_content, prf_output_path)

    # Surface OAuth accounts as an explicit, unmissable manual step rather than
    # letting them quietly vanish from the .prf.
    _, manual = partition_accounts(profile)
    for account in manual:
        manifest.notes.append(
            f"MANUAL STEP: account '{account.account_name}' ({account.email_address}) uses "
            f"OAuth2/modern auth and is NOT in the .prf. After importing, add it in Outlook via "
            f"File > Add Account and complete the provider's browser sign-in."
        )

    # Restore signatures/dictionary alongside, if present in the bundle.
    sig_src = bundle_dir / "signatures"
    if sig_src.exists():
        notes_target = target_pst_dir.parent / "restored_signatures"
        if notes_target.exists():
            shutil.rmtree(notes_target)
        shutil.copytree(sig_src, notes_target)
        manifest.notes.append(
            f"Signatures restored to {notes_target} — copy into %APPDATA%\\Microsoft\\Signatures "
            f"manually or via GUI helper."
        )

    return prf_output_path


def run_importprf(prf_path: Path, outlook_exe: str = "outlook.exe") -> subprocess.CompletedProcess:
    """
    Invokes `outlook.exe /importprf <prf_path>`. Separated from restore_profile()
    so callers (GUI/CLI) explicitly opt into actually launching Outlook.
    """
    if not prf_path.exists():
        raise RestoreError(f"PRF file not found: {prf_path}")
    try:
        return subprocess.run([outlook_exe, "/importprf", str(prf_path)], check=True)
    except FileNotFoundError as exc:
        raise RestoreError(
            f"Could not launch Outlook: '{outlook_exe}' not found. Pass --outlook-exe with the "
            f"full path (e.g. C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE)."
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise RestoreError(
            f"Outlook exited with code {exc.returncode} while importing the .prf. "
            f"The profile may be partially created — check Control Panel > Mail > Show Profiles."
        ) from exc
