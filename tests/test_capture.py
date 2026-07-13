import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest

from core.manifest import PstFile, sha256_file
from core.capture import (
    _copy_pst, _is_file_locked, capture_signatures, capture_dictionary,
    CaptureError,
)


def test_copy_pst_success(tmp_path):
    src = tmp_path / "leon.pst"
    content = b"pst file contents here"
    src.write_bytes(content)

    pst = PstFile(
        original_path=str(src),
        display_name="leon",
        size_bytes=len(content),
        sha256=sha256_file(src),
    )

    dest_dir = tmp_path / "bundle" / "psts"
    notes = []
    result_name = _copy_pst(pst, dest_dir, notes)

    assert (dest_dir / result_name).exists()
    assert (dest_dir / result_name).read_bytes() == content
    assert result_name.startswith("leon__")


def test_copy_pst_missing_source_raises(tmp_path):
    pst = PstFile(
        original_path=str(tmp_path / "does_not_exist.pst"),
        display_name="leon",
        size_bytes=0,
        sha256="0" * 64,
    )
    with pytest.raises(CaptureError):
        _copy_pst(pst, tmp_path / "bundle", manifest_notes=[])


def test_copy_pst_hash_mismatch_aborts_and_cleans_up(tmp_path):
    src = tmp_path / "leon.pst"
    src.write_bytes(b"actual content")

    # Deliberately wrong hash to simulate the source file having changed
    # since scan.py hashed it (e.g. Outlook briefly reopened it).
    pst = PstFile(
        original_path=str(src),
        display_name="leon",
        size_bytes=src.stat().st_size,
        sha256="f" * 64,
    )
    dest_dir = tmp_path / "bundle" / "psts"

    with pytest.raises(CaptureError):
        _copy_pst(pst, dest_dir, manifest_notes=[])

    # The partial/incorrect copy must not be left behind.
    leftover = list(dest_dir.glob("leon__*.pst")) if dest_dir.exists() else []
    assert leftover == []


def test_is_file_locked_false_for_normal_file(tmp_path):
    f = tmp_path / "normal.pst"
    f.write_bytes(b"data")
    assert _is_file_locked(f) is False


def test_capture_signatures_missing_dir_notes_and_returns_false(tmp_path):
    bundle = tmp_path / "bundle"
    appdata = tmp_path / "appdata_missing"
    notes = []
    result = capture_signatures(bundle, appdata, notes)
    assert result is False
    assert len(notes) == 1


def test_capture_signatures_copies_existing_dir(tmp_path):
    bundle = tmp_path / "bundle"
    appdata = tmp_path / "appdata"
    sig_dir = appdata / "Microsoft" / "Signatures"
    sig_dir.mkdir(parents=True)
    (sig_dir / "MySig.htm").write_text("<html>sig</html>")

    notes = []
    result = capture_signatures(bundle, appdata, notes)
    assert result is True
    assert (bundle / "signatures" / "MySig.htm").exists()


def test_capture_dictionary_no_dic_files_notes_and_returns_false(tmp_path):
    bundle = tmp_path / "bundle"
    uproof = tmp_path / "uproof"
    uproof.mkdir()
    notes = []
    result = capture_dictionary(bundle, uproof, notes)
    assert result is False
    assert len(notes) == 1


def test_capture_dictionary_copies_dic_files(tmp_path):
    bundle = tmp_path / "bundle"
    uproof = tmp_path / "uproof"
    uproof.mkdir()
    (uproof / "custom.dic").write_text("word1\nword2\n")

    notes = []
    result = capture_dictionary(bundle, uproof, notes)
    assert result is True
    assert (bundle / "dictionary" / "custom.dic").exists()
