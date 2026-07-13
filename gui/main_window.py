"""
gui/main_window.py — Outlook Porter GUI

Thin PyQt6 wrapper around core/. All the actual scan/capture/restore/verify
logic lives in core/ and is unit-tested there; this file only handles
presentation and threading so the UI never blocks during a hash or PST copy.

Run with:  python main.py --gui
(or)       python -m gui.main_window
"""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from typing import Callable, Any

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QLineEdit, QFileDialog, QTreeWidget,
    QTreeWidgetItem, QPlainTextEdit, QCheckBox, QComboBox, QGroupBox,
    QMessageBox, QSplitter, QInputDialog, QFrame,
)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.manifest import Manifest, OutlookProfile
from core.scan import scan_all_profiles, ScanError
from core.capture import capture_profile, CaptureError
from core.restore import restore_profile, run_importprf, RestoreError
from core.verify import run_restore_verify, VerifyError


# ---------------------------------------------------------------------------
# Background worker — runs a single blocking callable off the UI thread.
# ---------------------------------------------------------------------------
class Worker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, fn: Callable[[], Any]):
        super().__init__()
        self.fn = fn

    def run(self) -> None:
        try:
            result = self.fn()
        except Exception as exc:  # noqa: BLE001 - surfaced to user, not swallowed
            self.failed.emit(f"{exc}\n\n{traceback.format_exc()}")
            return
        self.finished.emit(result)


def run_in_background(parent: QObject, fn: Callable[[], Any],
                       on_done: Callable[[Any], None],
                       on_error: Callable[[str], None]) -> None:
    """Spins up a QThread for fn(), wires signals, keeps thread/worker alive on parent."""
    thread = QThread(parent)
    worker = Worker(fn)
    worker.moveToThread(thread)

    thread.started.connect(worker.run)
    worker.finished.connect(on_done)
    worker.failed.connect(on_error)
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    worker.failed.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)

    # keep references so Python doesn't GC them mid-run
    parent._thread = thread
    parent._worker = worker
    thread.start()


def find_outlook_exe() -> str:
    """
    Best-effort auto-locate outlook.exe so the Quick Migrate flow doesn't need
    to ask for it. Falls back to the bare command (relies on PATH) if nothing
    is found — restore_profile/run_importprf will surface a clear error if
    that guess is wrong, rather than silently doing nothing (fail closed).
    """
    candidates = []
    for env_var in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        base = os.environ.get(env_var)
        if not base:
            continue
        for office_dir in ("Office16", "Office15", "Office14"):
            candidates.append(Path(base) / "Microsoft Office" / "root" / office_dir / "OUTLOOK.EXE")
            candidates.append(Path(base) / "Microsoft Office" / office_dir / "OUTLOOK.EXE")
    for c in candidates:
        if c.exists():
            return str(c)
    return "outlook.exe"


def default_pst_dir() -> Path:
    """Best-effort default target folder for restored PSTs."""
    return Path.home() / "Documents" / "Outlook Files"


# ---------------------------------------------------------------------------
# Quick Migrate tab — the two-button front door. Wraps scan+capture into one
# action and restore+verify into another, using sane defaults so a technician
# doesn't need to touch the Advanced tabs for a normal job. Advanced tabs
# remain available for inspecting the manifest/prf or handling edge cases.
# ---------------------------------------------------------------------------
class QuickMigrateTab(QWidget):
    def __init__(self, log: Callable[[str], None]):
        super().__init__()
        self.log = log
        self._verify_bundle_dir: Path | None = None
        self._verify_pst_dir: Path | None = None

        layout = QVBoxLayout(self)

        # --- Backup section --------------------------------------------
        backup_group = QGroupBox("Step 1 — On the OLD PC")
        backup_layout = QVBoxLayout(backup_group)
        backup_layout.addWidget(QLabel(
            "Close Outlook completely first (check the system tray too).\n"
            "This scans, then captures PST + signatures + dictionary + rules in one go."
        ))
        self.backup_btn = QPushButton("Back Up This PC's Outlook...")
        self.backup_btn.clicked.connect(self.do_backup)
        backup_layout.addWidget(self.backup_btn)
        layout.addWidget(backup_group)

        # --- Restore section ---------------------------------------------
        restore_group = QGroupBox("Step 2 — On the NEW PC")
        restore_layout = QVBoxLayout(restore_group)
        restore_layout.addWidget(QLabel(
            "Copy the backup folder to this PC first (USB, network share, etc.), "
            "then restore. Outlook will launch automatically and prompt for the "
            "email password once — that's expected."
        ))
        self.restore_btn = QPushButton("Restore Outlook Here...")
        self.restore_btn.clicked.connect(self.do_restore)
        restore_layout.addWidget(self.restore_btn)

        self.verify_btn = QPushButton("I've signed in — Verify now")
        self.verify_btn.setEnabled(False)
        self.verify_btn.clicked.connect(self.do_verify)
        restore_layout.addWidget(self.verify_btn)

        layout.addWidget(restore_group)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        layout.addStretch()

    # ---- Backup ---------------------------------------------------------
    def do_backup(self) -> None:
        bundle_dir = QFileDialog.getExistingDirectory(self, "Choose a folder to save the backup into")
        if not bundle_dir:
            return
        bundle_dir = Path(bundle_dir)

        self.backup_btn.setEnabled(False)
        self.status_label.setText("")
        self.log("Scanning Outlook profiles...")

        scan_notes: list[str] = []

        def work():
            return scan_all_profiles(notes=scan_notes)

        def scan_done(profiles):
            for note in scan_notes:
                self.log(f"  note: {note}")
            if not profiles:
                self.backup_btn.setEnabled(True)
                self.log("SCAN: no POP3/IMAP profiles found.")
                QMessageBox.warning(self, "Nothing found",
                                     "No POP3/IMAP Outlook profiles were found on this PC.")
                return

            profile = profiles[0]
            if len(profiles) > 1:
                names = [p.profile_name for p in profiles]
                choice, ok = QInputDialog.getItem(
                    self, "Choose profile", "Multiple Outlook profiles found:", names, 0, False
                )
                if not ok:
                    self.backup_btn.setEnabled(True)
                    return
                profile = next(p for p in profiles if p.profile_name == choice)

            self.log(f"Capturing profile '{profile.profile_name}' -> {bundle_dir} ...")
            appdata = Path(os.environ.get("APPDATA", ""))
            uproof = appdata / "Microsoft" / "UProof"

            def capture_work():
                manifest = capture_profile(profile, bundle_dir, appdata, uproof)
                manifest.notes = scan_notes + manifest.notes
                manifest.save(bundle_dir / "manifest.json")
                return manifest

            def capture_done(manifest):
                self.backup_btn.setEnabled(True)
                self.log(f"Backup complete: {bundle_dir / 'manifest.json'}")
                rules_flag = False
                for note in manifest.notes:
                    self.log(f"  note: {note}")
                    if "rule" in note.lower():
                        rules_flag = True

                msg = f"Backup saved to:\n{bundle_dir}\n\nCopy this whole folder to the new PC."
                if rules_flag:
                    msg += (
                        "\n\nNote: check Outlook's own Manage Rules & Alerts on this PC — "
                        "if any rules are listed there, use its Export Rules (.rwz) option "
                        "and drop the file into the backup folder too (see log for details)."
                    )
                self.status_label.setText("Backup complete — see log for details.")
                QMessageBox.information(self, "Backup complete", msg)

            def capture_error(msg):
                self.backup_btn.setEnabled(True)
                self.log(f"CAPTURE FAILED: {msg}")
                QMessageBox.critical(self, "Backup failed", msg.splitlines()[0])

            run_in_background(self, capture_work, capture_done, capture_error)

        def scan_error(msg):
            self.backup_btn.setEnabled(True)
            self.log(f"SCAN FAILED: {msg}")
            QMessageBox.critical(self, "Backup failed", msg.splitlines()[0])

        run_in_background(self, work, scan_done, scan_error)

    # ---- Restore ----------------------------------------------------------
    def do_restore(self) -> None:
        bundle_dir = QFileDialog.getExistingDirectory(self, "Choose the backup folder to restore from")
        if not bundle_dir:
            return
        bundle_dir = Path(bundle_dir)
        if not (bundle_dir / "manifest.json").exists():
            QMessageBox.warning(self, "Not a backup folder",
                                 f"No manifest.json found in:\n{bundle_dir}\n\n"
                                 "Choose the folder created by Step 1 on the old PC.")
            return

        default_target = default_pst_dir()
        default_target.mkdir(parents=True, exist_ok=True)
        pst_dir = QFileDialog.getExistingDirectory(
            self, "Choose where to place restored PST files", str(default_target)
        )
        if not pst_dir:
            return
        pst_dir = Path(pst_dir)

        outlook_exe = find_outlook_exe()
        prf_path = bundle_dir / "restore.prf"

        self.restore_btn.setEnabled(False)
        self.verify_btn.setEnabled(False)
        self.status_label.setText("")
        self.log(f"Restoring {bundle_dir} -> {pst_dir} (outlook: {outlook_exe}) ...")

        def work():
            manifest = Manifest.load(bundle_dir / "manifest.json")
            restore_profile(manifest, bundle_dir, pst_dir, prf_path)
            run_importprf(prf_path, outlook_exe=outlook_exe)
            return manifest

        def done(manifest):
            self.restore_btn.setEnabled(True)
            self._verify_bundle_dir = bundle_dir
            self._verify_pst_dir = pst_dir
            self.verify_btn.setEnabled(True)
            self.log("Restore complete, Outlook launched with /importprf.")
            self.status_label.setText(
                "Outlook is launching and will ask for the email password once. "
                "After signing in and letting mail sync briefly, click "
                "'I've signed in — Verify now' below."
            )

        def error(msg):
            self.restore_btn.setEnabled(True)
            self.log(f"RESTORE FAILED: {msg}")
            QMessageBox.critical(
                self, "Restore failed",
                msg.splitlines()[0] + f"\n\nIf outlook.exe wasn't found automatically, "
                f"use the Advanced > Restore tab to specify its path directly."
            )

        run_in_background(self, work, done, error)

    # ---- Verify -------------------------------------------------------
    def do_verify(self) -> None:
        if not self._verify_bundle_dir or not self._verify_pst_dir:
            return
        bundle_dir = self._verify_bundle_dir
        pst_dir = self._verify_pst_dir

        self.verify_btn.setEnabled(False)
        self.log("Running restore-verify...")

        def work():
            manifest = Manifest.load(bundle_dir / "manifest.json")
            restored_paths = {}
            for profile in manifest.profiles:
                for account in profile.accounts:
                    if account.pst:
                        restored_paths[account.pst.original_path] = pst_dir / f"{account.pst.display_name}.pst"
            return run_restore_verify(manifest, restored_paths)

        def done(discrepancies):
            self.verify_btn.setEnabled(True)
            if discrepancies:
                self.status_label.setStyleSheet("color: #b00000; font-weight: bold;")
                self.status_label.setText("Discrepancies found:\n" + "\n".join(discrepancies))
                self.log("VERIFY: discrepancies found — see panel.")
            else:
                self.status_label.setStyleSheet("color: #007000; font-weight: bold;")
                self.status_label.setText("Migration complete — restored data matches the backup. Done.")
                self.log("VERIFY: clean pass. Migration complete.")

        def error(msg):
            self.verify_btn.setEnabled(True)
            self.log(f"VERIFY FAILED: {msg}")
            QMessageBox.critical(self, "Verify failed", msg.splitlines()[0])

        run_in_background(self, work, done, error)


# ---------------------------------------------------------------------------
# Scan tab
# ---------------------------------------------------------------------------
class ScanTab(QWidget):
    def __init__(self, log: Callable[[str], None]):
        super().__init__()
        self.log = log
        self.profiles: list[OutlookProfile] = []

        layout = QVBoxLayout(self)

        row = QHBoxLayout()
        self.scan_btn = QPushButton("Scan this PC")
        self.scan_btn.clicked.connect(self.do_scan)
        row.addWidget(self.scan_btn)
        row.addStretch()
        layout.addLayout(row)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Profile / Account", "Type", "Server", "PST"])
        layout.addWidget(self.tree)

        note = QLabel(
            "Scope: POP3 / IMAP accounts only. Exchange/M365 accounts are skipped "
            "by design — re-add those on the new PC and let them resync."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(note)

    def do_scan(self) -> None:
        self.scan_btn.setEnabled(False)
        self.log("Scanning Outlook profiles...")

        def work():
            return scan_all_profiles()

        def done(profiles: list[OutlookProfile]):
            self.scan_btn.setEnabled(True)
            self.profiles = profiles
            self.tree.clear()
            for profile in profiles:
                p_item = QTreeWidgetItem([f"{profile.profile_name}", "", f"Outlook {profile.outlook_version}", ""])
                self.tree.addTopLevelItem(p_item)
                for acct in profile.accounts:
                    pst_label = acct.pst.original_path if acct.pst else "(no PST)"
                    a_item = QTreeWidgetItem([
                        f"  {acct.account_name} <{acct.email_address}>",
                        acct.account_type,
                        f"{acct.incoming_server}:{acct.incoming_port}",
                        pst_label,
                    ])
                    p_item.addChild(a_item)
                p_item.setExpanded(True)
            self.log(f"Scan complete: {len(profiles)} profile(s) found.")

        def error(msg: str):
            self.scan_btn.setEnabled(True)
            self.log(f"SCAN FAILED: {msg}")
            QMessageBox.critical(self, "Scan failed", msg.splitlines()[0])

        run_in_background(self, work, done, error)


# ---------------------------------------------------------------------------
# Capture tab
# ---------------------------------------------------------------------------
class CaptureTab(QWidget):
    def __init__(self, log: Callable[[str], None], scan_tab: ScanTab):
        super().__init__()
        self.log = log
        self.scan_tab = scan_tab

        layout = QVBoxLayout(self)

        profile_row = QHBoxLayout()
        profile_row.addWidget(QLabel("Profile:"))
        self.profile_combo = QComboBox()
        profile_row.addWidget(self.profile_combo)
        refresh_btn = QPushButton("Refresh from scan")
        refresh_btn.clicked.connect(self.refresh_profiles)
        profile_row.addWidget(refresh_btn)
        profile_row.addStretch()
        layout.addLayout(profile_row)

        bundle_row = QHBoxLayout()
        bundle_row.addWidget(QLabel("Bundle output folder:"))
        self.bundle_edit = QLineEdit()
        bundle_row.addWidget(self.bundle_edit)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_bundle)
        bundle_row.addWidget(browse_btn)
        layout.addLayout(bundle_row)

        warn = QLabel("Close Outlook completely before capturing (check the system tray too).")
        warn.setStyleSheet("color: #b06000; font-weight: bold;")
        layout.addWidget(warn)

        self.capture_btn = QPushButton("Capture")
        self.capture_btn.clicked.connect(self.do_capture)
        layout.addWidget(self.capture_btn)

        layout.addStretch()

    def refresh_profiles(self) -> None:
        self.profile_combo.clear()
        for p in self.scan_tab.profiles:
            self.profile_combo.addItem(p.profile_name)
        if not self.scan_tab.profiles:
            self.log("No scanned profiles yet — run Scan first.")

    def browse_bundle(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose bundle output folder")
        if path:
            self.bundle_edit.setText(path)

    def do_capture(self) -> None:
        if not self.scan_tab.profiles:
            QMessageBox.warning(self, "No profile", "Run Scan first, then choose a profile here.")
            return
        if not self.bundle_edit.text():
            QMessageBox.warning(self, "No bundle folder", "Choose a bundle output folder first.")
            return

        profile_name = self.profile_combo.currentText()
        profile = next((p for p in self.scan_tab.profiles if p.profile_name == profile_name), None)
        if profile is None:
            QMessageBox.warning(self, "Profile not found", "Refresh from scan and pick a profile.")
            return

        bundle_dir = Path(self.bundle_edit.text())
        bundle_dir.mkdir(parents=True, exist_ok=True)

        import os
        appdata = Path(os.environ.get("APPDATA", ""))
        uproof = appdata / "Microsoft" / "UProof"

        self.capture_btn.setEnabled(False)
        self.log(f"Capturing profile '{profile_name}' -> {bundle_dir} ...")

        def work():
            manifest = capture_profile(profile, bundle_dir, appdata, uproof)
            manifest.save(bundle_dir / "manifest.json")
            return manifest

        def done(manifest: Manifest):
            self.capture_btn.setEnabled(True)
            self.log(f"Capture complete: {bundle_dir / 'manifest.json'}")
            for note in manifest.notes:
                self.log(f"  note: {note}")
            QMessageBox.information(self, "Capture complete", f"Bundle written to:\n{bundle_dir}")

        def error(msg: str):
            self.capture_btn.setEnabled(True)
            self.log(f"CAPTURE FAILED: {msg}")
            QMessageBox.critical(self, "Capture failed", msg.splitlines()[0])

        run_in_background(self, work, done, error)


# ---------------------------------------------------------------------------
# Restore tab
# ---------------------------------------------------------------------------
class RestoreTab(QWidget):
    def __init__(self, log: Callable[[str], None]):
        super().__init__()
        self.log = log

        layout = QVBoxLayout(self)

        bundle_row = QHBoxLayout()
        bundle_row.addWidget(QLabel("Bundle folder (from capture):"))
        self.bundle_edit = QLineEdit()
        bundle_row.addWidget(self.bundle_edit)
        b_browse = QPushButton("Browse...")
        b_browse.clicked.connect(lambda: self._browse(self.bundle_edit, files=False))
        bundle_row.addWidget(b_browse)
        layout.addLayout(bundle_row)

        pst_row = QHBoxLayout()
        pst_row.addWidget(QLabel("Destination PST folder:"))
        self.pst_edit = QLineEdit()
        pst_row.addWidget(self.pst_edit)
        p_browse = QPushButton("Browse...")
        p_browse.clicked.connect(lambda: self._browse(self.pst_edit, files=False))
        pst_row.addWidget(p_browse)
        layout.addLayout(pst_row)

        outlook_row = QHBoxLayout()
        outlook_row.addWidget(QLabel("outlook.exe path:"))
        self.outlook_edit = QLineEdit("outlook.exe")
        outlook_row.addWidget(self.outlook_edit)
        layout.addLayout(outlook_row)

        self.launch_checkbox = QCheckBox("Launch Outlook with /importprf automatically when done")
        layout.addWidget(self.launch_checkbox)

        cred_note = QLabel(
            "Passwords are never captured or restored. The client will be prompted "
            "to sign in once, on first send/receive — this is expected."
        )
        cred_note.setWordWrap(True)
        cred_note.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(cred_note)

        self.restore_btn = QPushButton("Restore")
        self.restore_btn.clicked.connect(self.do_restore)
        layout.addWidget(self.restore_btn)

        layout.addStretch()

    def _browse(self, target: QLineEdit, files: bool) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose folder")
        if path:
            target.setText(path)

    def do_restore(self) -> None:
        if not self.bundle_edit.text() or not self.pst_edit.text():
            QMessageBox.warning(self, "Missing paths", "Choose both the bundle folder and destination PST folder.")
            return

        bundle_dir = Path(self.bundle_edit.text())
        pst_dir = Path(self.pst_edit.text())
        prf_path = bundle_dir / "restore.prf"
        launch = self.launch_checkbox.isChecked()
        outlook_exe = self.outlook_edit.text() or "outlook.exe"

        self.restore_btn.setEnabled(False)
        self.log(f"Restoring bundle {bundle_dir} -> {pst_dir} ...")

        def work():
            manifest = Manifest.load(bundle_dir / "manifest.json")
            restore_profile(manifest, bundle_dir, pst_dir, prf_path)
            if launch:
                run_importprf(prf_path, outlook_exe=outlook_exe)
            return prf_path

        def done(result_prf: Path):
            self.restore_btn.setEnabled(True)
            self.log(f"Restore complete. PRF written to {result_prf}")
            if launch:
                self.log("Outlook launched with /importprf.")
                QMessageBox.information(self, "Restore complete", "PSTs restored and Outlook launched.")
            else:
                QMessageBox.information(
                    self, "Restore complete",
                    f"PSTs restored.\n\nRun manually:\n\"{self.outlook_edit.text()}\" /importprf \"{result_prf}\""
                )

        def error(msg: str):
            self.restore_btn.setEnabled(True)
            self.log(f"RESTORE FAILED: {msg}")
            QMessageBox.critical(self, "Restore failed", msg.splitlines()[0])

        run_in_background(self, work, done, error)


# ---------------------------------------------------------------------------
# Verify tab
# ---------------------------------------------------------------------------
class VerifyTab(QWidget):
    def __init__(self, log: Callable[[str], None]):
        super().__init__()
        self.log = log

        layout = QVBoxLayout(self)

        bundle_row = QHBoxLayout()
        bundle_row.addWidget(QLabel("Bundle folder:"))
        self.bundle_edit = QLineEdit()
        bundle_row.addWidget(self.bundle_edit)
        b_browse = QPushButton("Browse...")
        b_browse.clicked.connect(lambda: self._browse(self.bundle_edit))
        bundle_row.addWidget(b_browse)
        layout.addLayout(bundle_row)

        pst_row = QHBoxLayout()
        pst_row.addWidget(QLabel("Restored PST folder:"))
        self.pst_edit = QLineEdit()
        pst_row.addWidget(self.pst_edit)
        p_browse = QPushButton("Browse...")
        p_browse.clicked.connect(lambda: self._browse(self.pst_edit))
        pst_row.addWidget(p_browse)
        layout.addLayout(pst_row)

        self.verify_btn = QPushButton("Run restore-verify")
        self.verify_btn.clicked.connect(self.do_verify)
        layout.addWidget(self.verify_btn)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

        layout.addStretch()

    def _browse(self, target: QLineEdit) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose folder")
        if path:
            target.setText(path)

    def do_verify(self) -> None:
        if not self.bundle_edit.text() or not self.pst_edit.text():
            QMessageBox.warning(self, "Missing paths", "Choose both the bundle folder and restored PST folder.")
            return

        bundle_dir = Path(self.bundle_edit.text())
        pst_dir = Path(self.pst_edit.text())

        self.verify_btn.setEnabled(False)
        self.result_label.setText("Running...")
        self.log(f"Verifying restore against {bundle_dir} ...")

        def work():
            manifest = Manifest.load(bundle_dir / "manifest.json")
            restored_paths = {}
            for profile in manifest.profiles:
                for account in profile.accounts:
                    if account.pst:
                        restored_paths[account.pst.original_path] = pst_dir / f"{account.pst.display_name}.pst"
            return run_restore_verify(manifest, restored_paths)

        def done(discrepancies: list[str]):
            self.verify_btn.setEnabled(True)
            if discrepancies:
                self.result_label.setStyleSheet("color: #b00000; font-weight: bold;")
                self.result_label.setText("Discrepancies found:\n" + "\n".join(discrepancies))
                self.log("VERIFY: discrepancies found — see panel.")
            else:
                self.result_label.setStyleSheet("color: #007000; font-weight: bold;")
                self.result_label.setText("Clean pass — restored data matches captured baseline.")
                self.log("VERIFY: clean pass.")

        def error(msg: str):
            self.verify_btn.setEnabled(True)
            self.result_label.setText("")
            self.log(f"VERIFY FAILED: {msg}")
            QMessageBox.critical(self, "Verify failed", msg.splitlines()[0])

        run_in_background(self, work, done, error)


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outlook Porter — POP3/IMAP migration")
        self.resize(900, 650)

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)

        splitter = QSplitter(Qt.Orientation.Vertical)
        outer.addWidget(splitter)

        self.tabs = QTabWidget()
        splitter.addWidget(self.tabs)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)
        log_layout.addWidget(self.log_view)
        splitter.addWidget(log_group)
        splitter.setSizes([450, 150])

        quick_tab = QuickMigrateTab(self.log)
        self.tabs.addTab(quick_tab, "Quick Migrate")

        scan_tab = ScanTab(self.log)
        capture_tab = CaptureTab(self.log, scan_tab)
        restore_tab = RestoreTab(self.log)
        verify_tab = VerifyTab(self.log)

        advanced_tabs = QTabWidget()
        advanced_tabs.addTab(scan_tab, "1. Scan")
        advanced_tabs.addTab(capture_tab, "2. Capture")
        advanced_tabs.addTab(restore_tab, "3. Restore")
        advanced_tabs.addTab(verify_tab, "4. Verify")
        self.tabs.addTab(advanced_tabs, "Advanced")

    def log(self, message: str) -> None:
        self.log_view.appendPlainText(message)


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
