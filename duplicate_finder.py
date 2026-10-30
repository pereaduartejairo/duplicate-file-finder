#!/usr/bin/env python
"""Duplicate File Finder v2.2.

Scan a folder or volume for files with identical content, review the matches,
and move the copies you no longer need to the Recycle Bin.
"""

from __future__ import annotations

import ctypes
import hashlib
import os
import sys
from datetime import datetime

from PySide6.QtCore import QMutex, QObject, QThread, Qt, QWaitCondition, Signal, Slot
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

# ══════════════════════════════════════════════════════════════════════════
# Configuration
# ══════════════════════════════════════════════════════════════════════════

APP_NAME = "Duplicate File Finder"
APP_VERSION = "2.2.0"

# Directory names that are never descended into. Matching is done per folder
# name (not against the whole path), so a user folder such as
# "C:\Users\me\Documents\system-notes" is still scanned normally.
EXCLUDED_DIR_NAMES = {
    "program files",
    "program files (x86)",
    "windows",
    "system32",
    "system",
    "$recycle.bin",
    "system volume information",
    "recovery",
    "windows.old",
}

# ══════════════════════════════════════════════════════════════════════════
# Small helpers
# ══════════════════════════════════════════════════════════════════════════


def human_size(num_bytes: float) -> str:
    """Return a compact, human-readable size string."""
    size = float(num_bytes)
    if size <= 0:
        return "0 B"
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            if unit == "B":
                return f"{size:.0f} B"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


# ══════════════════════════════════════════════════════════════════════════
# Volume detection
# ══════════════════════════════════════════════════════════════════════════


def get_all_volumes() -> dict[str, str]:
    """Return a mapping of drive root -> volume label.

    Includes internal and external drives. Example:
    ``{"C:\\": "OS", "D:\\": "Backup"}``.
    """
    volumes: dict[str, str] = {}

    if sys.platform != "win32":
        volumes["/"] = "Root"
        return volumes

    import string

    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for index, letter in enumerate(string.ascii_uppercase):
        if not bitmask & (1 << index):
            continue

        drive = f"{letter}:\\"
        label = ""
        try:
            volume_name = ctypes.create_unicode_buffer(261)
            filesystem_name = ctypes.create_unicode_buffer(261)
            ok = ctypes.windll.kernel32.GetVolumeInformationW(
                drive,
                volume_name,
                len(volume_name),
                None,
                None,
                None,
                filesystem_name,
                len(filesystem_name),
            )
            if ok:
                label = volume_name.value.strip()
        except Exception:
            label = ""

        volumes[drive] = label or f"Drive {letter}:"

    return volumes


def get_root_drives() -> list[str]:
    """Return drive roots that currently exist and can be traversed."""
    drives: list[str] = []
    for drive in get_all_volumes():
        try:
            if os.path.exists(drive):
                drives.append(drive)
        except OSError:
            continue
    return drives


# ══════════════════════════════════════════════════════════════════════════
# File hashing
# ══════════════════════════════════════════════════════════════════════════


def compute_file_hash(
    filepath: str,
    chunk_size: int = 1024 * 1024,
    checkpoint=None,
) -> str:
    """Compute a SHA-256 digest, reading in chunks.

    ``checkpoint`` is called before each chunk. It may block (pause) and
    returns ``False`` to abort the hash (cancel). An empty string is returned
    when the file cannot be read or the hash was aborted.
    """
    try:
        hasher = hashlib.sha256()
        with open(filepath, "rb") as handle:
            while True:
                if checkpoint is not None and not checkpoint():
                    return ""
                chunk = handle.read(chunk_size)
                if not chunk:
                    break
                hasher.update(chunk)
        return hasher.hexdigest()
    except (OSError, PermissionError):
        return ""


# ══════════════════════════════════════════════════════════════════════════
# Recycle Bin support
# ══════════════════════════════════════════════════════════════════════════


class _SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("wFunc", ctypes.c_uint),
        ("pFrom", ctypes.c_wchar_p),
        ("pTo", ctypes.c_wchar_p),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", ctypes.c_int),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", ctypes.c_wchar_p),
    ]


def _recycle_with_shell_api(filepath: str) -> bool:
    """Move a file to the Recycle Bin using SHFileOperationW.

    This is a dependency-free fallback for systems where the pywin32
    Shell.Application approach is unavailable.
    """
    if sys.platform != "win32":
        return False

    FO_DELETE = 3
    FOF_SILENT = 0x0004
    FOF_NOCONFIRMATION = 0x0010
    FOF_ALLOWUNDO = 0x0040
    FOF_NOERRORUI = 0x0400

    try:
        operation = _SHFILEOPSTRUCTW()
        operation.hwnd = None
        operation.wFunc = FO_DELETE
        operation.pFrom = filepath + "\0"
        operation.pTo = None
        operation.fFlags = (
            FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
        )
        result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(operation))
        return result == 0 and not operation.fAnyOperationsAborted
    except Exception:
        return False


def send_to_recycle_bin(filepath: str) -> bool:
    """Move a file to the Recycle Bin.

    Never permanently deletes a file: if every available Recycle Bin method
    fails, the function returns ``False`` so the caller can report it.
    """
    path = os.path.abspath(filepath)
    if not os.path.exists(path):
        return False

    if sys.platform != "win32":
        try:
            from send2trash import send2trash

            send2trash(path)
            return True
        except Exception:
            return False

    # Preferred method: Shell.Application through pywin32.
    try:
        import win32com.client

        shell = win32com.client.Dispatch("Shell.Application")
        recycle_bin = shell.Namespace(10)
        item = shell.Namespace(os.path.dirname(path)).ParseName(
            os.path.basename(path)
        )
        if item is not None:
            recycle_bin.MoveHere(item)
            if not os.path.exists(path):
                return True
    except Exception:
        pass

    # Dependency-free Windows fallback.
    return _recycle_with_shell_api(path)

# ══════════════════════════════════════════════════════════════════════════
# Scan worker
# ══════════════════════════════════════════════════════════════════════════


class ScanWorker(QObject):
    """Find duplicate files in background threads.

    The scan uses two passes:
    1. Group files by size (cheap stat calls only).
    2. Hash only files that share a size with another file.

    Pause/resume and cancellation are coordinated with a mutex and a wait
    condition, so no polling loop or sleep is needed.
    """

    progress = Signal(int, int)  # current, total (total == 0 -> indeterminate)
    finished = Signal(dict)      # digest -> [paths]
    error = Signal(str)
    cancelled = Signal()

    def __init__(self, roots: list[str], excluded_names: set[str] | None = None):
        super().__init__()
        self.roots = [root for root in roots if root]
        self.excluded_names = {
            name.lower() for name in (excluded_names or EXCLUDED_DIR_NAMES)
        }
        self._mutex = QMutex()
        self._wait = QWaitCondition()
        self._cancelled = False
        self._paused = False

    # ── Control methods (called from the GUI thread) ───────────────────

    def pause(self) -> None:
        self._mutex.lock()
        self._paused = True
        self._mutex.unlock()

    def resume(self) -> None:
        self._mutex.lock()
        self._paused = False
        self._mutex.unlock()
        self._wait.wakeAll()

    def cancel(self) -> None:
        self._mutex.lock()
        self._cancelled = True
        self._paused = False
        self._mutex.unlock()
        self._wait.wakeAll()

    def _checkpoint(self) -> bool:
        """Block while paused; return False when the scan was cancelled."""
        self._mutex.lock()
        try:
            if self._cancelled:
                return False
            while self._paused and not self._cancelled:
                self._wait.wait(self._mutex)
            return not self._cancelled
        finally:
            self._mutex.unlock()

    @staticmethod
    def _is_reparse_point(path: str) -> bool:
        """Avoid following junctions/symlinks that could create loops."""
        try:
            if os.path.islink(path):
                return True
            is_junction = getattr(os.path, "isjunction", None)
            return bool(is_junction and is_junction(path))
        except OSError:
            return False

    # ── Worker entry point ─────────────────────────────────────────────

    @Slot()
    def run(self) -> None:
        try:
            size_groups = self._collect_by_size()
            if size_groups is None:
                self.cancelled.emit()
                return

            duplicates = self._hash_candidates(size_groups)
            if duplicates is None:
                self.cancelled.emit()
                return

            self.finished.emit(duplicates)
        except Exception as exc:  # keep the UI informed instead of crashing
            self.error.emit(f"{type(exc).__name__}: {exc}")

    # ── Pass 1: group by size ─────────────────────────────────────────

    def _collect_by_size(self) -> dict[int, list[str]] | None:
        size_groups: dict[int, list[str]] = {}
        scanned = 0

        for root in self.roots:
            if not self._checkpoint():
                return None

            for current, dirs, files in os.walk(
                root, topdown=True, onerror=lambda _error: None
            ):
                if not self._checkpoint():
                    return None

                dirs[:] = [
                    name
                    for name in dirs
                    if name.lower() not in self.excluded_names
                    and not self._is_reparse_point(os.path.join(current, name))
                ]

                for filename in files:
                    if not self._checkpoint():
                        return None

                    filepath = os.path.join(current, filename)
                    try:
                        file_size = os.path.getsize(filepath)
                    except OSError:
                        continue

                    scanned += 1
                    if file_size > 0:
                        size_groups.setdefault(file_size, []).append(filepath)

                    if scanned % 50 == 0:
                        self.progress.emit(scanned, 0)

        self.progress.emit(scanned, 0)
        return size_groups

    # ── Pass 2: hash only size collisions ─────────────────────────────

    def _hash_candidates(
        self, size_groups: dict[int, list[str]]
    ) -> dict[str, list[str]] | None:
        candidates = [
            paths for paths in size_groups.values() if len(paths) > 1
        ]
        total = sum(len(paths) for paths in candidates)

        if total == 0:
            self.progress.emit(0, 0)
            return {}

        hash_groups: dict[str, list[str]] = {}
        hashed = 0
        self.progress.emit(0, total)

        for paths in candidates:
            for filepath in paths:
                if not self._checkpoint():
                    return None

                digest = compute_file_hash(
                    filepath, checkpoint=self._checkpoint
                )
                if not self._checkpoint():
                    return None

                hashed += 1
                if digest:
                    hash_groups.setdefault(digest, []).append(filepath)

                if hashed % 10 == 0 or hashed == total:
                    self.progress.emit(hashed, total)

        return {
            digest: paths
            for digest, paths in hash_groups.items()
            if len(paths) > 1
        }

# ══════════════════════════════════════════════════════════════════════════
# UI helpers
# ══════════════════════════════════════════════════════════════════════════


def make_card() -> QFrame:
    """Create a white rounded card container."""
    card = QFrame()
    card.setObjectName("Card")
    card.setStyleSheet(
        "QFrame#Card { background: #ffffff; border: 1px solid #e5e7eb; "
        "border-radius: 10px; }"
    )
    return card


_PRIMARY_BUTTON_STYLE = """
QPushButton {
    background-color: #2563eb;
    color: #ffffff;
    border: none;
    border-radius: 7px;
    padding: 8px 16px;
    font-weight: 600;
}
QPushButton:hover { background-color: #1d4ed8; }
QPushButton:pressed { background-color: #1e40af; }
QPushButton:disabled { background-color: #e5e7eb; color: #9ca3af; }
"""

_SECONDARY_BUTTON_STYLE = """
QPushButton {
    background-color: #f9fafb;
    color: #1f2937;
    border: 1px solid #d1d5db;
    border-radius: 7px;
    padding: 8px 14px;
}
QPushButton:hover { background-color: #f3f4f6; border-color: #9ca3af; }
QPushButton:pressed { background-color: #e5e7eb; }
QPushButton:disabled {
    background-color: #f3f4f6;
    color: #9ca3af;
    border-color: #e5e7eb;
}
"""

_DANGER_BUTTON_STYLE = """
QPushButton {
    background-color: #dc2626;
    color: #ffffff;
    border: none;
    border-radius: 7px;
    padding: 8px 14px;
    font-weight: 600;
}
QPushButton:hover { background-color: #b91c1c; }
QPushButton:pressed { background-color: #991b1b; }
QPushButton:disabled { background-color: #e5e7eb; color: #9ca3af; }
"""


class ModernButton(QPushButton):
    """Push button with a consistent primary/secondary/danger style."""

    def __init__(
        self,
        text: str,
        primary: bool = False,
        danger: bool = False,
        parent: QWidget | None = None,
    ):
        super().__init__(text, parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(38)
        self.setFont(QFont("Segoe UI", 10))

        if primary:
            self.setStyleSheet(_PRIMARY_BUTTON_STYLE)
        elif danger:
            self.setStyleSheet(_DANGER_BUTTON_STYLE)
        else:
            self.setStyleSheet(_SECONDARY_BUTTON_STYLE)


# ══════════════════════════════════════════════════════════════════════════
# Main window
# ══════════════════════════════════════════════════════════════════════════


class DuplicateFinderApp(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"🔍 {APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(980, 720)
        self.resize(1180, 820)

        # State (created before the UI so slots can safely reference it).
        self.all_volumes: dict[str, str] = get_all_volumes()
        self.duplicate_groups: dict[str, list[str]] = {}
        self.selected_files: set[str] = set()
        self.scan_thread: QThread | None = None
        self.scan_worker: ScanWorker | None = None
        self._scanning = False
        self._paused = False
        self._busy = False
        self._last_scan_target = ""
        self._last_browse_dir = os.path.expanduser("~")

        self._build_ui()
        self._update_target_controls()
        self._update_stats()
        self._set_results_placeholder(
            "Choose a folder or volume, then click “Scan for duplicates”."
        )

    # ── Construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(22, 22, 22, 22)
        root.setSpacing(12)

        root.addWidget(self._build_title_card())
        root.addWidget(self._build_target_card())
        root.addWidget(self._build_stats_card())
        root.addWidget(self._build_results_card(), 1)
        root.addWidget(self._build_status_area())

    def _build_title_card(self) -> QFrame:
        card = make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(3)

        header = QHBoxLayout()
        title = QLabel("🔍 Duplicate File Finder")
        title.setFont(QFont("Segoe UI", 21, QFont.Weight.Bold))
        title.setStyleSheet("color: #111827;")

        version = QLabel(f"v{APP_VERSION}")
        version.setStyleSheet(
            "background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; "
            "border-radius: 10px; padding: 3px 10px; font-size: 10px; "
            "font-weight: 600;"
        )

        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(version)

        subtitle = QLabel(
            "Find identical files by content, review the matches, and move "
            "the extras to the Recycle Bin."
        )
        subtitle.setFont(QFont("Segoe UI", 10))
        subtitle.setStyleSheet("color: #6b7280;")

        layout.addLayout(header)
        layout.addWidget(subtitle)
        return card

    def _build_target_card(self) -> QFrame:
        card = make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(9)

        caption = QLabel("1 · Choose what to scan")
        caption.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        caption.setStyleSheet("color: #374151;")
        layout.addWidget(caption)

        # Folder row ---------------------------------------------------
        folder_row = QHBoxLayout()
        folder_row.setSpacing(8)

        folder_label = QLabel("Folder")
        folder_label.setFixedWidth(52)
        folder_label.setFont(QFont("Segoe UI", 10))
        folder_label.setStyleSheet("color: #6b7280;")

        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText(
            "Pick a folder with Browse, or type/paste a path…"
        )
        self.folder_path.setClearButtonEnabled(True)
        self.folder_path.setMinimumHeight(36)
        self.folder_path.setStyleSheet(
            """
            QLineEdit {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 10px;
                color: #1f2937;
                font-size: 11px;
            }
            QLineEdit:focus { border-color: #2563eb; }
            QLineEdit:read-only { background: #f9fafb; color: #6b7280; }
            """
        )
        self.folder_path.textChanged.connect(self._on_folder_changed)

        self.browse_btn = ModernButton("📁 Browse…")
        self.browse_btn.setMinimumWidth(105)
        self.browse_btn.setToolTip("Choose the folder you want to scan")
        self.browse_btn.clicked.connect(self.browse_folder)

        self.clear_btn = ModernButton("✖ Clear")
        self.clear_btn.setMinimumWidth(88)
        self.clear_btn.setToolTip(
            "Clear the folder and go back to scanning a volume"
        )
        self.clear_btn.clicked.connect(self.clear_folder)

        folder_row.addWidget(folder_label)
        folder_row.addWidget(self.folder_path, 1)
        folder_row.addWidget(self.browse_btn)
        folder_row.addWidget(self.clear_btn)
        layout.addLayout(folder_row)

        # Volume row ---------------------------------------------------
        volume_row = QHBoxLayout()
        volume_row.setSpacing(8)

        volume_label = QLabel("Volume")
        volume_label.setFixedWidth(52)
        volume_label.setFont(QFont("Segoe UI", 10))
        volume_label.setStyleSheet("color: #6b7280;")

        self.volume_combo = QComboBox()
        self.volume_combo.setMinimumHeight(36)
        self.volume_combo.setMinimumWidth(300)
        self.volume_combo.setStyleSheet(
            """
            QComboBox {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 10px;
                color: #1f2937;
                font-size: 11px;
            }
            QComboBox:hover { border-color: #9ca3af; }
            QComboBox:disabled { background: #f3f4f6; color: #9ca3af; }
            """
        )
        self._populate_volumes()
        self.volume_combo.currentIndexChanged.connect(self._on_volume_changed)

        volume_hint = QLabel("Used only when the folder field is empty.")
        volume_hint.setFont(QFont("Segoe UI", 9))
        volume_hint.setStyleSheet("color: #9ca3af;")

        volume_row.addWidget(volume_label)
        volume_row.addWidget(self.volume_combo)
        volume_row.addWidget(volume_hint)
        volume_row.addStretch(1)
        layout.addLayout(volume_row)

        # Action row ---------------------------------------------------
        action_row = QHBoxLayout()
        action_row.setSpacing(8)

        self.scan_btn = ModernButton("🔍 Scan for duplicates", primary=True)
        self.scan_btn.setMinimumWidth(180)
        self.scan_btn.clicked.connect(self.start_scan)

        self.pause_btn = ModernButton("⏸ Pause")
        self.pause_btn.clicked.connect(self.pause_scan)
        self.pause_btn.setEnabled(False)
        self.pause_btn.setToolTip("Pause the scan at the next checkpoint")

        self.continue_btn = ModernButton("▶ Continue")
        self.continue_btn.clicked.connect(self.resume_scan)
        self.continue_btn.setEnabled(False)
        self.continue_btn.setToolTip("Resume a paused scan")

        self.cancel_btn = ModernButton("✖ Cancel")
        self.cancel_btn.clicked.connect(self.cancel_scan)
        self.cancel_btn.setEnabled(False)

        self.target_info = QLabel("Ready")
        self.target_info.setFont(QFont("Segoe UI", 9))
        self.target_info.setStyleSheet("color: #6b7280;")
        self.target_info.setToolTip("Current scan target")

        action_row.addWidget(self.scan_btn)
        action_row.addWidget(self.pause_btn)
        action_row.addWidget(self.continue_btn)
        action_row.addWidget(self.cancel_btn)
        action_row.addStretch(1)
        action_row.addWidget(self.target_info)
        layout.addLayout(action_row)

        return card

    def _populate_volumes(self) -> None:
        """Fill the volume combo box with all logical drives."""
        self.volume_combo.clear()
        self.volume_combo.addItem("All drives", userData="all")

        for drive, label in sorted(self.all_volumes.items()):
            drive_name = drive.rstrip("\\/") or drive
            # Avoid "Drive C: (C:)" when Windows has no volume label.
            if label.startswith("Drive "):
                display = label
            else:
                display = f"{label} ({drive_name})"
            self.volume_combo.addItem(display, userData=drive)

        self.volume_combo.setCurrentIndex(0)

    def _build_stats_card(self) -> QFrame:
        card = make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(8)

        caption = QLabel("2 · Review and clean up")
        caption.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        caption.setStyleSheet("color: #374151;")
        layout.addWidget(caption)

        row = QHBoxLayout()
        row.setSpacing(18)

        self.stats_labels: dict[str, QLabel] = {}
        for key, title_text in (
            ("files", "Duplicate files"),
            ("groups", "Groups"),
            ("space", "Reclaimable"),
        ):
            box = QWidget()
            box_layout = QVBoxLayout(box)
            box_layout.setContentsMargins(0, 0, 0, 0)
            box_layout.setSpacing(0)

            title = QLabel(title_text)
            title.setFont(QFont("Segoe UI", 9))
            title.setStyleSheet("color: #6b7280;")

            value = QLabel("0")
            value.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
            value.setStyleSheet("color: #111827;")

            self.stats_labels[key] = value
            box_layout.addWidget(title)
            box_layout.addWidget(value)
            box.setMinimumWidth(120)
            row.addWidget(box)

        row.addStretch(1)

        self.selection_label = QLabel("No files selected")
        self.selection_label.setFont(QFont("Segoe UI", 9))
        self.selection_label.setStyleSheet("color: #6b7280;")

        self.report_btn = ModernButton("📄 Report")
        self.report_btn.clicked.connect(self.generate_report)
        self.report_btn.setEnabled(False)
        self.report_btn.setToolTip("Save the current results as a text report")

        self.delete_btn = ModernButton("🗑 Move to Recycle Bin", danger=True)
        self.delete_btn.clicked.connect(self.delete_selected)
        self.delete_btn.setEnabled(False)

        row.addWidget(self.selection_label)
        row.addWidget(self.report_btn)
        row.addWidget(self.delete_btn)
        layout.addLayout(row)

        return card

    def _build_results_card(self) -> QFrame:
        card = make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("Duplicate groups")
        title.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        title.setStyleSheet("color: #374151;")

        self.results_label = QLabel("No results yet")
        self.results_label.setFont(QFont("Segoe UI", 9))
        self.results_label.setStyleSheet("color: #9ca3af;")

        header.addWidget(title)
        header.addWidget(self.results_label)
        header.addStretch(1)
        layout.addLayout(header)

        hint = QLabel(
            "Check each copy you want to remove. Keep at least one copy in "
            "every group."
        )
        hint.setFont(QFont("Segoe UI", 9))
        hint.setStyleSheet("color: #9ca3af;")
        layout.addWidget(hint)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.scroll_area.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.scroll_area.setStyleSheet(
            "QScrollArea { border: none; background: transparent; }"
        )

        self.scroll_content = QWidget()
        self.scroll_content.setStyleSheet("background: transparent;")
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(8)

        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area, 1)
        return card

    def _build_status_area(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(4, 0, 4, 0)
        layout.setSpacing(4)

        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Segoe UI", 9))
        self.status_label.setStyleSheet("color: #6b7280;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(
            """
            QProgressBar {
                border: none;
                background: #e5e7eb;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background: #2563eb;
                border-radius: 4px;
            }
            """
        )

        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)
        return container

    # ── Target selection ──────────────────────────────────────────────

    def _on_folder_changed(self, _text: str) -> None:
        self._update_target_controls()
        if not self._is_scanning():
            self.target_info.setText(self._mode_text())

    def _on_volume_changed(self, _index: int) -> None:
        if not self._is_scanning():
            self.target_info.setText(self._mode_text())

    def _mode_text(self) -> str:
        folder = self.folder_path.text().strip()
        if folder:
            normalized = os.path.normpath(folder)
            name = os.path.basename(normalized) or normalized
            return f"Folder: {name}"

        data = self.volume_combo.currentData()
        if data in (None, "all"):
            return "All drives"

        name = str(data).rstrip("\\/") or str(data)
        return f"Volume: {name}"

    def _update_target_controls(self) -> None:
        """Keep target controls usable and in sync with the scan state."""
        enabled = not self._is_scanning() and not self._busy
        has_folder = bool(self.folder_path.text().strip())

        self.folder_path.setReadOnly(not enabled)
        self.browse_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled and has_folder)
        self.volume_combo.setEnabled(enabled and not has_folder)

        if not self._is_scanning():
            if has_folder:
                self.target_info.setToolTip(self.folder_path.text().strip())
            else:
                data = self.volume_combo.currentData()
                self.target_info.setToolTip(
                    "All drives"
                    if data in (None, "all")
                    else str(data)
                )

    def _set_scanning_ui(self, scanning: bool, target: str = "") -> None:
        if scanning:
            self.scan_btn.setEnabled(False)
            self.pause_btn.setEnabled(True)
            self.continue_btn.setEnabled(False)
            self.cancel_btn.setEnabled(True)
            self.delete_btn.setEnabled(False)
            self.report_btn.setEnabled(False)
            self.target_info.setText("Scanning…")
            self.target_info.setToolTip(target)
        else:
            self.pause_btn.setEnabled(False)
            self.continue_btn.setEnabled(False)
            self.cancel_btn.setEnabled(False)
            self.scan_btn.setEnabled(
                not self._busy and not self._is_scanning()
            )
            self.report_btn.setEnabled(
                bool(self.duplicate_groups) and not self._busy
            )
            self._update_selection_label()
            self.target_info.setText(self._mode_text())

        self._update_target_controls()

    def browse_folder(self) -> None:
        """Open a folder picker and apply the selected folder.

        Uses the instance API with fully scoped enums (reliable on PySide6
        6.x) and falls back to the simple static call if anything goes
        wrong. Cancelling the dialog simply leaves the current target alone.
        """
        if self._is_scanning() or self._busy:
            return

        current = self.folder_path.text().strip().strip('"').strip("'")
        start_dir = current if os.path.isdir(current) else self._last_browse_dir
        if not os.path.isdir(start_dir):
            start_dir = os.path.expanduser("~")

        try:
            dialog = QFileDialog(self, "Select a folder to scan", start_dir)
            dialog.setFileMode(QFileDialog.FileMode.Directory)
            dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)
            dialog.setOption(QFileDialog.Option.DontResolveSymlinks, True)
            dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
            if dialog.exec():
                selected = dialog.selectedFiles()
                if selected:
                    self._apply_folder(selected[0])
            return
        except Exception:
            # Fall through to the dependency-free static API.
            pass

        try:
            folder = QFileDialog.getExistingDirectory(
                self, "Select a folder to scan", start_dir
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Folder picker unavailable",
                f"Could not open the folder picker:\n{exc}",
            )
            return

        if folder:
            self._apply_folder(folder)

    def _apply_folder(self, folder: str) -> None:
        path = os.path.normpath(folder)
        self.folder_path.setText(path)
        if os.path.isdir(path):
            self._last_browse_dir = path
        self.status_label.setText(f"Folder selected: {path}")

    def clear_folder(self) -> None:
        """Clear the folder so the volume selector can be used again."""
        if self._is_scanning() or self._busy:
            return
        self.folder_path.clear()
        self.status_label.setText(
            "Folder cleared — volume selection is active."
        )

    def _is_scanning(self) -> bool:
        if self._scanning:
            return True
        thread = self.scan_thread
        return bool(thread is not None and thread.isRunning())

    # ── Scan control ──────────────────────────────────────────────────

    def start_scan(self) -> None:
        """Validate the target and start a background scan."""
        if self._is_scanning() or self._busy:
            return

        folder_text = self.folder_path.text().strip().strip('"').strip("'")

        # "C:" means "current directory on C:" to the OS; treat it as C:\.
        if sys.platform == "win32" and len(folder_text) == 2 and folder_text[1] == ":":
            folder_text += "\\"

        if folder_text:
            folder = os.path.abspath(
                os.path.expanduser(os.path.expandvars(folder_text))
            )
            if not os.path.isdir(folder):
                QMessageBox.warning(
                    self,
                    "Folder not found",
                    "This folder does not exist or is not accessible:\n\n"
                    f"{folder}",
                )
                return
            roots = [folder]
            target = f"Folder: {folder}"
            self.folder_path.setText(folder)
        else:
            data = self.volume_combo.currentData()
            if data in (None, "all"):
                roots = get_root_drives()
                if not roots:
                    QMessageBox.warning(
                        self,
                        "No drives found",
                        "No readable drives were found to scan.",
                    )
                    return
                target = "All drives: " + ", ".join(
                    root.rstrip("\\/") for root in roots
                )
            else:
                drive = str(data)
                if not os.path.exists(drive):
                    QMessageBox.warning(
                        self,
                        "Drive unavailable",
                        f"This drive is not available:\n\n{drive}",
                    )
                    return
                roots = [drive]
                target = f"Drive: {drive.rstrip('\\/') or drive}"

        # Reset previous results.
        self.selected_files.clear()
        self.duplicate_groups = {}
        self._update_stats()
        self._set_results_placeholder("Scanning for duplicate files…")
        self._last_scan_target = target

        worker = ScanWorker(roots)
        thread = QThread(self)
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.progress.connect(self._on_scan_progress)
        worker.finished.connect(self._on_scan_finished)
        worker.error.connect(self._on_scan_error)
        worker.cancelled.connect(self._on_scan_cancelled)

        # Stop the thread event loop as soon as the worker is done.
        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        worker.cancelled.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(self._on_thread_finished)
        thread.finished.connect(thread.deleteLater)

        self.scan_worker = worker
        self.scan_thread = thread
        self._scanning = True
        self._paused = False

        self._set_scanning_ui(True, target)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("Scanning…")

        try:
            thread.start()
        except Exception as exc:
            self._scanning = False
            self._set_scanning_ui(False)
            self.progress_bar.setVisible(False)
            self.scan_thread = None
            self.scan_worker = None
            QMessageBox.critical(
                self, "Could not start scan", str(exc)
            )

    def pause_scan(self) -> None:
        if not self._is_scanning() or self.scan_worker is None or self._paused:
            return

        self.scan_worker.pause()
        self._paused = True
        self.pause_btn.setEnabled(False)
        self.continue_btn.setEnabled(True)
        self.status_label.setText("Scan paused — click Continue to resume.")
        self.target_info.setText("Paused")
        self.target_info.setToolTip(self._last_scan_target)

    def resume_scan(self) -> None:
        if not self._is_scanning() or self.scan_worker is None or not self._paused:
            return

        self.scan_worker.resume()
        self._paused = False
        self.pause_btn.setEnabled(True)
        self.continue_btn.setEnabled(False)
        self.status_label.setText("Scanning…")
        self.target_info.setText("Scanning…")
        self.target_info.setToolTip(self._last_scan_target)

    def cancel_scan(self) -> None:
        if not self._is_scanning() or self.scan_worker is None:
            return

        self.scan_worker.cancel()
        self._paused = False
        self.pause_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.status_label.setText("Cancelling scan…")
        self.target_info.setText("Cancelling…")

    # ── Worker signal handlers ────────────────────────────────────────

    def _on_scan_progress(self, current: int, total: int) -> None:
        if total <= 0:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
            self.status_label.setText(
                f"Scanning… {current:,} files checked"
            )
            return

        if self.progress_bar.maximum() != total:
            self.progress_bar.setRange(0, total)
        self.progress_bar.setValue(min(current, total))
        self.status_label.setText(
            f"Hashing candidate files… {current:,} of {total:,}"
        )

    def _on_scan_finished(self, duplicate_groups: dict) -> None:
        self._scanning = False
        self._paused = False
        self.duplicate_groups = dict(duplicate_groups)

        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self._update_stats()
        self._display_duplicate_groups()
        self._set_scanning_ui(False)

        group_count = len(self.duplicate_groups)
        total_files = sum(
            len(paths) for paths in self.duplicate_groups.values()
        )
        if group_count:
            self.status_label.setText(
                f"Scan complete — {group_count:,} group(s), "
                f"{total_files:,} duplicate file(s)."
            )
        else:
            self.status_label.setText(
                "Scan complete — no duplicates found."
            )

    def _on_scan_error(self, error_msg: str) -> None:
        self._scanning = False
        self._paused = False
        self.duplicate_groups = {}
        self.selected_files.clear()

        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self._update_stats()
        self._set_results_placeholder(
            "The scan failed. Check the message and try again."
        )
        self._set_scanning_ui(False)
        self.target_info.setText("Scan failed")
        self.status_label.setText("Scan failed")

        QMessageBox.critical(self, "Scan error", error_msg)

    def _on_scan_cancelled(self) -> None:
        self._scanning = False
        self._paused = False
        self.duplicate_groups = {}
        self.selected_files.clear()

        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self._update_stats()
        self._set_results_placeholder("Scan cancelled.")
        self._set_scanning_ui(False)
        self.target_info.setText("Scan cancelled")
        self.status_label.setText("Scan cancelled.")

    def _on_thread_finished(self) -> None:
        """Final cleanup after the QThread has really stopped."""
        thread = self.sender()
        if thread is not self.scan_thread:
            return

        self._scanning = False
        self._paused = False
        self.scan_thread = None
        self.scan_worker = None

        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self._update_target_controls()
        self.scan_btn.setEnabled(not self._busy)
        self.pause_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self._update_selection_label()
        self.report_btn.setEnabled(
            bool(self.duplicate_groups) and not self._busy
        )

    # ── Statistics and selection ──────────────────────────────────────

    def _update_stats(self) -> None:
        group_count = len(self.duplicate_groups)
        total_files = sum(
            len(paths) for paths in self.duplicate_groups.values()
        )
        reclaimable = self._calculate_reclaimable_bytes()

        self.stats_labels["files"].setText(f"{total_files:,}")
        self.stats_labels["groups"].setText(f"{group_count:,}")
        self.stats_labels["space"].setText(human_size(reclaimable))

        if group_count:
            suffix = "s" if group_count != 1 else ""
            self.results_label.setText(
                f"{group_count:,} group{suffix}"
            )
        else:
            self.results_label.setText("No results yet")

        self._update_selection_label()

    def _calculate_reclaimable_bytes(self) -> int:
        """Bytes that would be freed by keeping one copy per group."""
        total = 0
        for paths in self.duplicate_groups.values():
            if len(paths) < 2:
                continue
            try:
                total += os.path.getsize(paths[0]) * (len(paths) - 1)
            except OSError:
                continue
        return total

    def _update_selection_label(self) -> None:
        count = len(self.selected_files)
        if count:
            suffix = "s" if count != 1 else ""
            self.selection_label.setText(
                f"{count:,} file{suffix} selected"
            )
        else:
            self.selection_label.setText("No files selected")

        self.delete_btn.setEnabled(
            count > 0 and not self._is_scanning() and not self._busy
        )

    # ── Deletion ──────────────────────────────────────────────────────

    def delete_selected(self) -> None:
        """Move the checked files to the Recycle Bin."""
        if self._is_scanning() or self._busy:
            return

        if not self.selected_files:
            QMessageBox.information(
                self,
                "No files selected",
                "Check one or more duplicate files first.",
            )
            return

        complete_groups = sum(
            1
            for paths in self.duplicate_groups.values()
            if all(path in self.selected_files for path in paths)
        )

        message = (
            f"Move {len(self.selected_files):,} selected file(s) to the "
            "Recycle Bin?"
        )
        if complete_groups:
            message += (
                f"\n\nWarning: every copy is selected in {complete_groups} "
                "group(s). Keep at least one copy if you may need these files."
            )

        reply = QMessageBox.question(
            self,
            "Confirm move to Recycle Bin",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self._busy = True
        self._update_target_controls()
        self.scan_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)
        self.report_btn.setEnabled(False)

        targets = sorted(self.selected_files)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(targets))
        self.progress_bar.setValue(0)

        succeeded: list[str] = []
        failed: list[str] = []

        for index, path in enumerate(targets, 1):
            if send_to_recycle_bin(path):
                succeeded.append(path)
            else:
                failed.append(path)

            self.progress_bar.setValue(index)
            self.status_label.setText(
                f"Moving to Recycle Bin… {index:,} of {len(targets):,}"
            )
            QApplication.processEvents()

        self._remove_deleted_paths(set(succeeded))
        self._busy = False

        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self._update_stats()
        self._display_duplicate_groups()
        self._set_scanning_ui(False)

        message = (
            f"Moved {len(succeeded):,} file(s) to the Recycle Bin"
            + (f" — {len(failed):,} failed" if failed else ".")
        )
        self.status_label.setText(message)

        if failed:
            preview = "\n".join(failed[:5])
            if len(failed) > 5:
                preview += f"\n… and {len(failed) - 5:,} more"
            QMessageBox.warning(
                self,
                "Some files were not moved",
                f"{len(failed):,} file(s) could not be moved to the "
                f"Recycle Bin:\n\n{preview}",
            )

    def _remove_deleted_paths(self, deleted: set[str]) -> None:
        """Drop moved files and any group that no longer has 2 copies."""
        updated: dict[str, list[str]] = {}
        for digest, paths in self.duplicate_groups.items():
            remaining = [path for path in paths if path not in deleted]
            if len(remaining) > 1:
                updated[digest] = remaining

        self.duplicate_groups = updated

        # Never keep a selection pointing at a group that is gone.
        visible = {
            path for paths in updated.values() for path in paths
        }
        self.selected_files &= visible

    # ── Report ────────────────────────────────────────────────────────

    def generate_report(self) -> None:
        """Write a readable text report of the current results."""
        if not self.duplicate_groups:
            QMessageBox.information(
                self,
                "No results",
                "Run a scan before generating a report.",
            )
            return

        default_name = (
            f"duplicate_report_{datetime.now():%Y%m%d_%H%M%S}.txt"
        )
        default_path = os.path.join(
            os.path.expanduser("~"), default_name
        )
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save duplicate report",
            default_path,
            "Text files (*.txt);;All files (*)",
        )
        if not save_path:
            return

        group_count = len(self.duplicate_groups)
        total_files = sum(
            len(paths) for paths in self.duplicate_groups.values()
        )
        reclaimable = self._calculate_reclaimable_bytes()

        lines: list[str] = [
            "=" * 72,
            f"{APP_NAME} v{APP_VERSION} — Duplicate report",
            "=" * 72,
            "",
            f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}",
            f"Scan target: {self._last_scan_target or 'unknown'}",
            "",
            "SUMMARY",
            "-" * 72,
            f"Duplicate groups : {group_count:,}",
            f"Duplicate files  : {total_files:,}",
            f"Reclaimable space: {human_size(reclaimable)}",
            f"Selected to remove: {len(self.selected_files):,} file(s)",
            "",
            "DETAILS",
            "-" * 72,
        ]

        sorted_groups = sorted(
            self.duplicate_groups.items(),
            key=lambda item: (-len(item[1]), item[1][0].lower()),
        )

        for index, (digest, paths) in enumerate(sorted_groups, 1):
            try:
                size = os.path.getsize(paths[0])
            except OSError:
                size = 0

            lines.extend(
                [
                    "",
                    f"Group {index} — {len(paths)} files, "
                    f"{human_size(size)} each",
                    f"  SHA-256: {digest}",
                ]
            )
            for path in paths:
                marker = "  [selected]" if path in self.selected_files else ""
                lines.append(f"  - {path}{marker}")

        try:
            with open(save_path, "w", encoding="utf-8") as handle:
                handle.write("\n".join(lines) + "\n")
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Report error",
                f"Could not save the report:\n{exc}",
            )
            return

        self.status_label.setText(f"Report saved: {save_path}")
        QMessageBox.information(
            self,
            "Report saved",
            f"The report was saved to:\n\n{save_path}",
        )

    # ── Results display ───────────────────────────────────────────────

    def _clear_results(self) -> None:
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _set_results_placeholder(self, text: str) -> None:
        self._clear_results()

        label = QLabel(text)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setWordWrap(True)
        label.setFont(QFont("Segoe UI", 11))
        label.setStyleSheet("color: #6b7280; padding: 32px;")

        self.scroll_layout.addWidget(label)
        self.scroll_layout.addStretch(1)

    def _display_duplicate_groups(self) -> None:
        if not self.duplicate_groups:
            self._set_results_placeholder("No duplicate files found.")
            return

        self._clear_results()

        sorted_groups = sorted(
            self.duplicate_groups.items(),
            key=lambda item: (-len(item[1]), item[1][0].lower()),
        )
        for index, (digest, paths) in enumerate(sorted_groups, 1):
            self._add_duplicate_group(index, digest, paths)

        self.scroll_layout.addStretch(1)

    def _add_duplicate_group(
        self, index: int, digest: str, paths: list[str]
    ) -> None:
        frame = QFrame()
        frame.setObjectName("GroupCard")
        frame.setStyleSheet(
            "#GroupCard { background: #f9fafb; border: 1px solid #e5e7eb; "
            "border-radius: 8px; }"
        )

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(4)

        header = QHBoxLayout()
        header.setSpacing(10)

        title = QLabel(
            f"Group {index} — {len(paths)} identical files"
        )
        title.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        title.setStyleSheet("color: #111827;")

        try:
            size = os.path.getsize(paths[0])
        except OSError:
            size = 0
        size_label = QLabel(
            f"{human_size(size)} each" if size else "Size unavailable"
        )
        size_label.setFont(QFont("Segoe UI", 9))
        size_label.setStyleSheet("color: #6b7280;")

        digest_label = QLabel(f"SHA-256 {digest[:12]}…")
        digest_label.setFont(QFont("Segoe UI", 9))
        digest_label.setStyleSheet("color: #9ca3af;")
        digest_label.setToolTip(digest)

        select_btn = QPushButton("Select all but first")
        select_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        select_btn.setStyleSheet(
            "QPushButton { border: none; background: transparent; "
            "color: #2563eb; font-size: 9px; padding: 2px 4px; } "
            "QPushButton:hover { color: #1d4ed8; }"
        )
        select_btn.setToolTip(
            "Check every copy except the first one in this group"
        )

        header.addWidget(title)
        header.addWidget(size_label)
        header.addStretch(1)
        header.addWidget(digest_label)
        header.addWidget(select_btn)
        layout.addLayout(header)

        checks: list[QCheckBox] = []
        for path in paths:
            row = QHBoxLayout()
            row.setSpacing(8)

            checkbox = QCheckBox()
            checkbox.setFixedWidth(22)
            checkbox.setStyleSheet(
                "QCheckBox::indicator { width: 16px; height: 16px; }"
            )
            checkbox.setChecked(path in self.selected_files)
            checkbox.toggled.connect(
                lambda checked, p=path: self._on_file_toggled(p, checked)
            )

            path_label = QLabel(path)
            path_label.setWordWrap(True)
            path_label.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
            )
            path_label.setFont(QFont("Segoe UI", 10))
            path_label.setStyleSheet("color: #374151;")
            path_label.setToolTip(path)

            row.addWidget(checkbox)
            row.addWidget(path_label, 1)
            layout.addLayout(row)
            checks.append(checkbox)

        def select_others() -> None:
            for check in checks[1:]:
                check.setChecked(True)

        select_btn.clicked.connect(select_others)
        self.scroll_layout.addWidget(frame)

    def _on_file_toggled(self, path: str, checked: bool) -> None:
        if checked:
            self.selected_files.add(path)
        else:
            self.selected_files.discard(path)
        self._update_selection_label()

    # ── Window lifecycle ──────────────────────────────────────────────

    def closeEvent(self, event) -> None:
        if self.scan_worker is not None:
            self.scan_worker.cancel()

        thread = self.scan_thread
        if thread is not None:
            thread.quit()
            if not thread.wait(5000):
                # Last resort on application exit; normal cancellation is
                # checked between 1 MB hash chunks, so this should not run.
                thread.terminate()
                thread.wait(1000)

        event.accept()


def main() -> int:
    """Application entry point."""
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("File Tools")
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))

    palette = app.palette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#f1f5f9"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#1f2937"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#f8fafc"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#1f2937"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#f3f4f6"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#1f2937"))
    app.setPalette(palette)

    window = DuplicateFinderApp()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
