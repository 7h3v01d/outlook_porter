import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest

from core.manifest import PstFile, VerificationCounts, sha256_file
from core.verify import verify_pst_hash, verify_counts_match, VerifyError


def test_verify_pst_hash_matches(tmp_path):
    f = tmp_path / "leon.pst"
    f.write_bytes(b"some pst content")
    pst = PstFile(
        original_path="C:\\old\\leon.pst",
        display_name="leon",
        size_bytes=f.stat().st_size,
        sha256=sha256_file(f),
    )
    verify_pst_hash(pst, f)  # should not raise


def test_verify_pst_hash_mismatch_raises(tmp_path):
    f = tmp_path / "leon.pst"
    f.write_bytes(b"some pst content")
    pst = PstFile(
        original_path="C:\\old\\leon.pst",
        display_name="leon",
        size_bytes=f.stat().st_size,
        sha256="0" * 64,
    )
    with pytest.raises(VerifyError):
        verify_pst_hash(pst, f)


def test_verify_pst_hash_missing_file_raises(tmp_path):
    pst = PstFile(original_path="x", display_name="leon", size_bytes=0, sha256="0" * 64)
    with pytest.raises(VerifyError):
        verify_pst_hash(pst, tmp_path / "does_not_exist.pst")


def test_verify_counts_match_identical():
    baseline = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100, "Sent Items": 20})
    current = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100, "Sent Items": 20})
    assert verify_counts_match(baseline, current) == []


def test_verify_counts_match_detects_mismatch():
    baseline = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100})
    current = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 95})
    discrepancies = verify_counts_match(baseline, current)
    assert len(discrepancies) == 1
    assert "Inbox" in discrepancies[0]
    assert "100" in discrepancies[0] and "95" in discrepancies[0]


def test_verify_counts_match_detects_missing_folder():
    baseline = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100, "Archive": 5})
    current = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100})
    discrepancies = verify_counts_match(baseline, current)
    assert any("Archive" in d and "missing" in d.lower() for d in discrepancies)


def test_verify_counts_match_detects_new_folder():
    baseline = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100})
    current = VerificationCounts(pst_display_name="leon", folder_counts={"Inbox": 100, "New Folder": 3})
    discrepancies = verify_counts_match(baseline, current)
    assert any("New Folder" in d and "appeared" in d.lower() for d in discrepancies)
