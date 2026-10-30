# Duplicate File Finder v2.2

A professional Windows desktop app for finding duplicate files by content, reviewing the matches, and moving the copies you no longer need to the Recycle Bin.

Built with **Python + PySide6**. The UI is split into clear sections: choose a target, scan, review per-file checkboxes, and clean up.

## ✨ What’s new in v2.2

### Folder selection fixed and simplified
- **Browse…** opens a real folder picker using PySide6’s scoped enums, with a safe fallback if the dialog cannot be created.
- The folder field is now **editable**, so you can also type or paste a path such as `D:\Backups\Photos`.
- **Clear** resets folder mode and immediately re-enables the volume selector.
- Browse and the volume selector are no longer left disabled after a scan, cancel, or error.
- The target controls were moved out of the old single cramped row, so **Browse can no longer be squeezed off-screen**.

### Pause and Continue
- **⏸ Pause** stops the worker at the next checkpoint (even in the middle of hashing a large file).
- **▶ Continue** resumes exactly where the scan stopped.
- Both buttons enable/disable automatically, and the status line shows whether the scan is running or paused.

### Review improvements
- **One checkbox per file**, not one per group. You choose exactly which copies to remove.
- **Select all but first** fills the checkboxes for a group while keeping one copy safe.
- Live selection count and a disabled delete button until something is checked.
- Clear warning if every copy in a group is selected.
- Reclaimable-space and duplicate-file totals use readable units (B / KB / MB / GB / TB).

### Safer cleanup
- Files go to the Windows **Recycle Bin**. The app never permanently deletes a file as a fallback.
- Recycle Bin support uses `Shell.Application` (pywin32) first and the native `SHFileOperationW` API second.
- Hashing checks for cancellation between 1 MB chunks, so Cancel and window close stay responsive.
- Worker threads are stopped and cleaned up correctly; starting a second scan cannot race the first.

### Other fixes
- Report generation now writes real line breaks and the correct per-file size.
- The report includes the scan target, full SHA-256 hashes, and a `[selected]` marker.
- Folder exclusion matches **folder names**, not substrings of the full path. A folder named `system-notes` is no longer skipped just because it contains “system”.
- Junctions and symlinks are not followed, avoiding traversal loops.
- The version is now **2.2.0** everywhere.

## 🛠️ Features

### Scanning
- Scan **all drives**, a **specific drive**, or a **specific folder**.
- Internal and external volumes are detected through the Windows API.
- System folders are excluded by name: `Program Files`, `Program Files (x86)`, `Windows`, `System32`, `System`, `$Recycle.Bin`, `System Volume Information`, `Recovery`, and `Windows.old`.
- Two-pass detection: group by size first, then SHA-256 only the size collisions.
- Background `QThread` keeps the UI responsive.
- Indeterminate progress while enumerating, then a determinate hashing progress bar.
- Pause, Continue, and Cancel.

### Review and cleanup
- Per-file checkboxes with full paths.
- “Select all but first” shortcut per group.
- Duplicate groups sorted by number of copies.
- Move selected files to the Recycle Bin with confirmation.
- Generate a detailed text report.

## 🚀 Installation

### Requirements
- Windows 10/11
- Python **3.10 or newer**
- PySide6

```bash
pip install PySide6
```

`pywin32` is optional but recommended for the best Recycle Bin integration:

```bash
pip install pywin32
```

Without pywin32, the app falls back to the Windows `SHFileOperationW` API.

### Run

```bash
cd C:\path\to\duplicate-file-finder
python duplicate_finder.py
```

## 🎯 How to use

### Scan a folder
1. Click **📁 Browse…** and choose a folder, or type/paste a path into the folder field.
2. Click **🔍 Scan for duplicates**.
3. Wait for the scan to finish.
4. Check each copy you want to remove. Use **Select all but first** to speed this up.
5. Click **🗑 Move to Recycle Bin** and confirm.
6. Click **📄 Report** to save the results.

### Scan a drive or all drives
1. Leave the folder field empty. (Use **Clear** if a folder is still selected.)
2. Choose **All drives** or a specific drive in the **Volume** dropdown.
3. Click **🔍 Scan for duplicates**.

### Pause and continue
1. Start a scan.
2. Click **⏸ Pause** — the worker blocks at its next checkpoint.
3. Click **▶ Continue** — the scan resumes from the same place.
4. Click **✖ Cancel** to stop the scan entirely.

### Change the target after a scan
- Click **Clear** to go back to volume mode, or **Browse…** to pick a different folder.
- The controls are always restored when a scan finishes, fails, or is cancelled.

## 📄 Report format

The report is a UTF-8 text file containing:

- Generation date/time and the exact scan target.
- Summary: groups, duplicate files, reclaimable space, selected files.
- Each group with its full SHA-256 hash, size per file, and every path.
- `[selected]` next to paths that were checked when the report was created.

## 🐛 Troubleshooting

**Browse does not open**
- The app uses a normal Windows folder dialog. If it is blocked by a policy or a remote session, type/paste the folder path directly into the folder field and click Scan.
- Run the app from a terminal (`python duplicate_finder.py`) to see any error message.

**A folder scan finds nothing**
- Make sure the folder contains at least two files with identical content and non-zero size.
- Empty files are ignored because removing them frees no space.
- The folder itself is scanned even if its name matches an excluded system folder; only subfolders with excluded names are skipped.

**A file cannot be moved**
- Close any program that has the file open and try again.
- Files on network shares may not have a Recycle Bin; the app reports the failure instead of deleting permanently.

## 🧱 Project structure

```text
duplicate_finder.py        Desktop app (Python + PySide6), v2.2
duplicate_finder_html.html Standalone browser demo
test_run.py                GUI smoke test
test_scan.py               Disposable-folder scan test
README.md                  This file
LICENSE                    MIT license
```

## 🛠️ Technology

- Python 3.10+
- PySide6 / Qt 6
- `hashlib` SHA-256
- `QThread`, `QMutex`, `QWaitCondition` for background scanning and pause/resume
- Windows Shell API / pywin32 for the Recycle Bin

## 📜 Version history

### v2.2.0 (current)
- Fixed folder selection and permanently-disabled target controls.
- Refactored the UI into target, action, stats, results, and status sections.
- Editable folder path, Clear button, working Browse on PySide6 6.x.
- Per-file checkboxes and “Select all but first”.
- Separate Pause and Continue buttons with real pause/resume.
- Safer Recycle Bin fallback; no permanent-delete fallback.
- Cancellable hashing, clean thread shutdown, report fixes.
- Folder exclusions now match folder names instead of path substrings.

### v2.1.0
- Added Pause/Continue and specific folder selection.
- Added folder path display and version 2.1.0.

### v2.0.0
- Volume selector with all drives and individual drive letters.
- External drive support and Windows API volume labels.
- Two-pass scanning optimization and modern UI.

### v1.0.0
- Initial duplicate finder with checkboxes, Recycle Bin cleanup, and reports.

## 📦 Build a standalone executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name DuplicateFinder duplicate_finder.py
```

The resulting executable is in `dist\DuplicateFinder.exe`.

## 📜 License

MIT — see [LICENSE](LICENSE).

---

**Duplicate File Finder v2.2.0** — choose a folder or volume, scan by content, review every copy, and clean up safely.
