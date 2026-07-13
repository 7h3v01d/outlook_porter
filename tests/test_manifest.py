import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest

from core.manifest import (
    Manifest, ManifestError, MailAccount, OutlookProfile, PstFile,
    new_manifest, sha256_file,
)


def make_sample_manifest() -> Manifest:
    m = new_manifest(source_hostname="TESTPC", windows_username="leon")
    pst = PstFile(
        original_path=r"C:\Users\leon\Documents\Outlook Files\leon.pst",
        display_name="leon",
        size_bytes=123456,
        sha256="a" * 64,
        is_default_delivery=True,
        captured_filename="leon__aaaaaaaaaa.pst",
    )
    account = MailAccount(
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
        pst=pst,
    )
    profile = OutlookProfile(
        profile_name="Outlook",
        is_default_profile=True,
        outlook_version="16.0",
        accounts=[account],
    )
    m.profiles.append(profile)
    return m


def test_round_trip(tmp_path):
    m = make_sample_manifest()
    path = tmp_path / "manifest.json"
    m.save(path)

    loaded = Manifest.load(path)
    assert loaded.source_hostname == "TESTPC"
    assert len(loaded.profiles) == 1
    assert loaded.profiles[0].accounts[0].email_address == "client@example.com"
    assert loaded.profiles[0].accounts[0].pst.sha256 == "a" * 64
    assert loaded.profiles[0].accounts[0].pst.captured_filename == "leon__aaaaaaaaaa.pst"


def test_load_missing_file_raises(tmp_path):
    with pytest.raises(ManifestError):
        Manifest.load(tmp_path / "does_not_exist.json")


def test_load_bad_json_raises(tmp_path):
    path = tmp_path / "manifest.json"
    path.write_text("{ not valid json", encoding="utf-8")
    with pytest.raises(ManifestError):
        Manifest.load(path)


def test_load_wrong_schema_version_raises_fail_closed(tmp_path):
    m = make_sample_manifest()
    path = tmp_path / "manifest.json"
    m.save(path)

    # Tamper with schema version to simulate a future/incompatible format
    import json
    data = json.loads(path.read_text())
    data["schema_version"] = 999
    path.write_text(json.dumps(data), encoding="utf-8")

    with pytest.raises(ManifestError):
        Manifest.load(path)


def test_account_without_pst_round_trips(tmp_path):
    m = new_manifest("HOST", "user")
    account = MailAccount(
        account_name="No PST Account",
        email_address="nopst@example.com",
        display_name="No PST",
        account_type="POP3",
        incoming_server="pop.example.com",
        incoming_port=995,
        incoming_ssl=True,
        outgoing_server="smtp.example.com",
        outgoing_port=465,
        outgoing_ssl=True,
        outgoing_auth=True,
        username="nopst@example.com",
        pst=None,
    )
    m.profiles.append(OutlookProfile(profile_name="P", is_default_profile=True, outlook_version="16.0", accounts=[account]))
    path = tmp_path / "m.json"
    m.save(path)
    loaded = Manifest.load(path)
    assert loaded.profiles[0].accounts[0].pst is None


def test_sha256_file(tmp_path):
    f = tmp_path / "sample.bin"
    f.write_bytes(b"hello world")
    import hashlib
    expected = hashlib.sha256(b"hello world").hexdigest()
    assert sha256_file(f) == expected
