import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest

from core.manifest import MailAccount, OutlookProfile, PstFile
from core.restore import build_prf, RestoreError, restore_profile, _restore_pst
from core.manifest import Manifest, new_manifest


def make_account(**overrides) -> MailAccount:
    defaults = dict(
        account_name="Client Mail",
        email_address="client@example.com",
        display_name="Client Name",
        account_type="IMAP",
        incoming_server="imap.example.com",
        incoming_port=993,
        incoming_ssl=True,
        outgoing_server="smtp.example.com",
        outgoing_port=587,
        outgoing_ssl=True,
        outgoing_auth=True,
        username="client@example.com",
        pst=None,
        outgoing_secure_connection=1,
    )
    defaults.update(overrides)
    return MailAccount(**defaults)


def test_build_prf_uses_real_schema_sections():
    """The .prf must use the documented 7-section format, not invented keys."""
    account = make_account()
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {})

    # Section 3 declares the account against a Section 7 mapping block.
    assert "[Internet Account List]" in prf
    assert "Account1=IMAP_I_Mail" in prf
    # Section 5 holds the values, under a plain [Account1] header.
    assert "[Account1]" in prf
    assert "EmailAddress=client@example.com" in prf
    assert "IMAPServer=imap.example.com" in prf
    assert "IMAPPort=993" in prf
    assert "IMAPUseSSL=1" in prf
    assert "SMTPServer=smtp.example.com" in prf
    # Sections 6 and 7 mapping blocks must be present or Outlook can't map properties.
    assert "[IMAP_I_Mail]" in prf
    assert "SMTPSecureConnection=PT_LONG,0x020A" in prf
    # The old invented key names must be gone.
    assert "Account1IMAPServer=" not in prf
    assert "Account1EmailAddress=" not in prf


def test_build_prf_emits_only_matching_account_type_keys():
    """An IMAP account must not also emit POP3Server= pointing at the IMAP host."""
    account = make_account(account_type="IMAP", incoming_server="imap.example.com")
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {})

    account_section = prf.split("[Account1]")[1].split("[")[0]
    assert "IMAPServer=imap.example.com" in account_section
    assert "POP3Server=" not in account_section


def test_build_prf_never_writes_a_password_value():
    """
    Section 6/7 legitimately contain the *mapping* line 'Password=PT_STRING8,...'.
    What must never appear is an actual password VALUE in sections 1-5.
    """
    account = make_account()
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {})

    settings = prf.split("; Section 6")[0]   # everything the technician configures
    assert "Password" not in settings


def test_starttls_account_emits_secure_connection_2():
    """The bug that broke sending: port 587 + STARTTLS must not become 'SSL=Yes'."""
    account = make_account(outgoing_port=587, outgoing_ssl=True, outgoing_secure_connection=2)
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {})

    assert "SMTPPort=587" in prf
    assert "SMTPSecureConnection=2" in prf


def test_implicit_ssl_account_emits_secure_connection_1():
    account = make_account(outgoing_port=465, outgoing_ssl=True, outgoing_secure_connection=1)
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {})

    assert "SMTPPort=465" in prf
    assert "SMTPSecureConnection=1" in prf


def test_oauth_account_is_excluded_from_prf():
    """A Gmail/OAuth2 account cannot be provisioned by .prf — it must not be emitted."""
    basic = make_account(account_name="Basic", email_address="basic@example.com")
    oauth = make_account(account_name="Gmail", email_address="user@gmail.com",
                         uses_oauth=True, auth_identity_uid="123_tp_google_imap_OAuth2")
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0",
                             accounts=[basic, oauth])
    prf = build_prf(profile, {})

    assert "basic@example.com" in prf
    assert "user@gmail.com" not in prf
    # only the one provisionable account is declared
    assert "Account1=IMAP_I_Mail" in prf
    assert "Account2=" not in prf


def test_all_oauth_profile_fails_closed():
    oauth = make_account(uses_oauth=True)
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[oauth])
    with pytest.raises(RestoreError, match="no accounts that can be provisioned"):
        build_prf(profile, {})


def test_restored_pst_is_bound_as_a_service():
    """build_prf previously ignored restored_pst_paths entirely."""
    pst = PstFile(original_path=r"C:\old\mail.pst", display_name="mail",
                  size_bytes=10, sha256="a" * 64, is_default_delivery=True,
                  captured_filename="mail__aaaaaaaaaa.pst")
    account = make_account(account_type="POP3", incoming_server="pop.example.com", pst=pst)
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])
    prf = build_prf(profile, {r"C:\old\mail.pst": Path(r"D:\new\mail.pst")})

    assert "Service1=Unicode Personal Folders" in prf
    assert "PathAndFilenameToPersonalFolders=" in prf
    assert "DefaultStore=Service1" in prf


def test_build_prf_multiple_accounts_increment_index():
    a1 = make_account(account_name="Acct1", email_address="one@example.com")
    a2 = make_account(account_name="Acct2", email_address="two@example.com")
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[a1, a2])
    prf = build_prf(profile, {})

    assert "Account1=IMAP_I_Mail" in prf
    assert "Account2=IMAP_I_Mail" in prf
    assert "[Account1]" in prf
    assert "[Account2]" in prf
    assert "EmailAddress=one@example.com" in prf
    assert "EmailAddress=two@example.com" in prf


def test_build_prf_no_accounts_raises():
    profile = OutlookProfile(profile_name="Empty", is_default_profile=True, outlook_version="16.0", accounts=[])
    with pytest.raises(RestoreError):
        build_prf(profile, {})


def test_restore_pst_refuses_to_overwrite(tmp_path):
    bundle_dir = tmp_path / "bundle"
    (bundle_dir / "psts").mkdir(parents=True)
    src_pst_content = b"fake pst bytes"
    (bundle_dir / "psts" / "leon__aaaa.pst").write_bytes(src_pst_content)

    target_dir = tmp_path / "target"
    target_dir.mkdir()
    (target_dir / "leon.pst").write_bytes(b"existing file")  # already occupies destination

    pst = PstFile(
        original_path=r"C:\old\leon.pst",
        display_name="leon",
        size_bytes=len(src_pst_content),
        sha256="dummy",
        captured_filename="leon__aaaa.pst",
    )

    with pytest.raises(RestoreError):
        _restore_pst(pst, bundle_dir, target_dir, notes=[])


def test_restore_pst_missing_captured_file_raises(tmp_path):
    bundle_dir = tmp_path / "bundle"
    bundle_dir.mkdir()
    target_dir = tmp_path / "target"

    pst = PstFile(
        original_path=r"C:\old\leon.pst",
        display_name="leon",
        size_bytes=10,
        sha256="dummy",
        captured_filename=None,
    )
    with pytest.raises(RestoreError):
        _restore_pst(pst, bundle_dir, target_dir, notes=[])


def test_restore_profile_end_to_end(tmp_path):
    import hashlib

    bundle_dir = tmp_path / "bundle"
    (bundle_dir / "psts").mkdir(parents=True)
    pst_bytes = b"pretend this is a pst file"
    (bundle_dir / "psts" / "leon__abc123.pst").write_bytes(pst_bytes)
    sha = hashlib.sha256(pst_bytes).hexdigest()

    pst = PstFile(
        original_path=r"C:\old\leon.pst",
        display_name="leon",
        size_bytes=len(pst_bytes),
        sha256=sha,
        captured_filename="leon__abc123.pst",
    )
    account = make_account(pst=pst)
    profile = OutlookProfile(profile_name="Outlook", is_default_profile=True, outlook_version="16.0", accounts=[account])

    manifest = new_manifest("HOST", "user")
    manifest.profiles.append(profile)

    target_dir = tmp_path / "target"
    prf_path = tmp_path / "restore.prf"

    result_prf_path = restore_profile(manifest, bundle_dir, target_dir, prf_path)

    assert result_prf_path == prf_path
    assert prf_path.exists()
    assert (target_dir / "leon.pst").read_bytes() == pst_bytes


def test_restore_profile_rejects_multi_profile_manifest(tmp_path):
    manifest = new_manifest("HOST", "user")
    manifest.profiles.append(OutlookProfile(profile_name="P1", is_default_profile=True, outlook_version="16.0", accounts=[make_account()]))
    manifest.profiles.append(OutlookProfile(profile_name="P2", is_default_profile=False, outlook_version="16.0", accounts=[make_account()]))

    with pytest.raises(RestoreError):
        restore_profile(manifest, tmp_path / "bundle", tmp_path / "target", tmp_path / "restore.prf")
