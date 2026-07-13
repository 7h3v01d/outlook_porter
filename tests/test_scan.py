import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import core.scan as scan_module
from core.scan import find_archive_psts


def test_find_archive_psts_finds_unclaimed_pst(tmp_path):
    search_dir = tmp_path / "Outlook"
    search_dir.mkdir()
    (search_dir / "archive.pst").write_bytes(b"pst bytes")

    notes: list[str] = []
    result = find_archive_psts(known_pst_paths=set(), search_dirs=[search_dir], notes=notes)

    assert len(result) == 1
    assert result[0].display_name == "archive"
    assert result[0].is_default_delivery is False
    assert any("Found 1 archive" in n for n in notes)


def test_find_archive_psts_skips_known_default_delivery_pst(tmp_path):
    search_dir = tmp_path / "Outlook"
    search_dir.mkdir()
    known = search_dir / "main.pst"
    known.write_bytes(b"pst bytes")

    result = find_archive_psts(known_pst_paths={str(known)}, search_dirs=[search_dir])

    assert result == []


def test_find_archive_psts_missing_dir_returns_empty(tmp_path):
    missing = tmp_path / "does_not_exist"
    result = find_archive_psts(known_pst_paths=set(), search_dirs=[missing])
    assert result == []


def test_find_archive_psts_two_dirs_no_duplicates(tmp_path):
    dir_a = tmp_path / "a"
    dir_b = tmp_path / "b"
    dir_a.mkdir()
    dir_b.mkdir()
    (dir_a / "archive.pst").write_bytes(b"same name, different dir")
    (dir_b / "archive.pst").write_bytes(b"same name, different dir 2")

    result = find_archive_psts(known_pst_paths=set(), search_dirs=[dir_a, dir_b])

    # Different absolute paths -> both should be captured, not deduped.
    assert len(result) == 2


# --- Regression test: mocked registry shaped like a real machine ----------
#
# This reproduces the exact structure pulled via `reg query` from a live
# Outlook profile: subkey 1 is the address book (no IMAP/POP3/Email fields,
# different clsid), subkey 2 is a real IMAP account whose "Delivery Store
# EntryID" points at an .ost (cache), not a .pst. There is no "Account Type"
# value anywhere — this is what caused every account to be silently dropped
# before the fix.

def _utf16le_entry_id(path: str) -> bytes:
    """Build a fake entry-id blob: some header bytes + provider name + UTF16LE path."""
    return b"\x00\x00\x00\x00pstprx.dll\x00\x00" + path.encode("utf-16-le") + b"\x00\x00"


class _FakeKey:
    def __init__(self, node):
        self.node = node

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class _FakeWinreg:
    HKEY_CURRENT_USER = object()

    def __init__(self, tree):
        self.tree = tree

    def OpenKey(self, parent, subpath=None):
        if parent is self.HKEY_CURRENT_USER:
            node = self.tree
            parts = subpath.split("\\")
        else:
            node = parent.node
            parts = [subpath]
        for part in parts:
            if part not in node.get("_subkeys", {}):
                raise FileNotFoundError(part)
            node = node["_subkeys"][part]
        return _FakeKey(node)

    def EnumKey(self, key, index):
        names = list(key.node.get("_subkeys", {}).keys())
        if index >= len(names):
            raise OSError("no more subkeys")
        return names[index]

    def QueryValueEx(self, key, name):
        values = key.node.get("_values", {})
        if name not in values:
            raise FileNotFoundError(name)
        return values[name]


def _make_fake_registry():
    address_book = {
        "_values": {
            "clsid": ("{ED475414-B0D6-11D2-8C3B-00104B2A6676}", 1),
            "Service Name": ("CONTAB", 1),
            "Account Name": ("Outlook Address Book", 1),
        }
    }
    imap_account = {
        "_values": {
            "clsid": ("{ED475412-B0D6-11D2-8C3B-00104B2A6676}", 1),
            "Account Name": ("ken.n.priest@bigpond.com", 1),
            "Email": ("ken.n.priest@bigpond.com", 1),
            "Display Name": ("ken.n.priest@bigpond.com", 1),
            "IMAP Server": ("imap.bigpond.com", 1),
            "IMAP User": ("ken.n.priest@bigpond.com", 1),
            "IMAP Port": (993, 4),
            "IMAP Use SSL": (1, 4),
            "SMTP Server": ("smtp.bigpond.com", 1),
            "SMTP Port": (465, 4),
            "SMTP Use SSL": (1, 4),
            "SMTP Authenticate": (1, 4),
            "Delivery Store EntryID": (
                _utf16le_entry_id(r"C:\Users\kennp\AppData\Local\Microsoft\Outlook\ken.n.priest@bigpond.com.ost"),
                3,
            ),
        }
    }
    guid_key = {"_subkeys": {"00000001": address_book, "00000002": imap_account}}
    profile_key = {"_subkeys": {scan_module.ACCOUNT_SUBKEY_GUID: guid_key}}
    profiles_key = {"_subkeys": {"Outlook": profile_key}}
    root = {"_subkeys": {"Software\\Microsoft\\Office\\16.0\\Outlook\\Profiles": profiles_key}}
    # Collapse the multi-part root path into the single-lookup form OpenKey expects
    # by flattening the profiles key directly under its full path string.
    tree = {"_subkeys": {}}
    tree["_subkeys"]["Software"] = {"_subkeys": {"Microsoft": {"_subkeys": {"Office": {"_subkeys": {"16.0": {"_subkeys": {"Outlook": {"_subkeys": {"Profiles": profiles_key}}}}}}}}}}
    return tree


def test_scan_profile_finds_real_imap_account_and_skips_address_book(monkeypatch):
    fake_tree = _make_fake_registry()
    fake_winreg = _FakeWinreg(fake_tree)
    monkeypatch.setattr(scan_module, "_open_winreg", lambda: fake_winreg)

    notes: list[str] = []
    profile = scan_module.scan_profile("16.0", "Outlook", notes=notes)

    assert len(profile.accounts) == 1
    account = profile.accounts[0]
    assert account.email_address == "ken.n.priest@bigpond.com"
    assert account.account_type == "IMAP"
    assert account.incoming_server == "imap.bigpond.com"
    # The account's store is an .ost (cache) — must NOT be treated as a PST to migrate.
    assert account.pst is None

    # Address book subkey must be skipped, not mistaken for an account.
    assert any("not a mail account" in n for n in notes)
    # The .ost should be explicitly noted as intentionally not captured.
    assert any(".ost" in n and "not captured" in n for n in notes)
