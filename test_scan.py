"""Functional scan test using a disposable folder."""

import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from PySide6.QtCore import QCoreApplication

from duplicate_finder import ScanWorker


def main() -> None:
    root = Path(tempfile.mkdtemp(prefix="dup_finder_test_"))
    try:
        payload = b"identical payload\n" * 10
        (root / "a").mkdir()
        (root / "b").mkdir()
        (root / "a" / "copy.txt").write_bytes(payload)
        (root / "b" / "copy.txt").write_bytes(payload)
        (root / "unique.txt").write_bytes(b"unique")
        (root / "empty1.txt").write_bytes(b"")
        (root / "empty2.txt").write_bytes(b"")

        app = QCoreApplication([])  # noqa: F841 - needed for Qt signals
        worker = ScanWorker([str(root)])

        groups: dict[str, list[str]] = {}
        errors: list[str] = []
        worker.finished.connect(lambda result: groups.update(result))
        worker.error.connect(errors.append)

        # Synchronous run keeps this test deterministic and fast.
        worker.run()

        assert not errors, errors
        assert len(groups) == 1, groups
        paths = next(iter(groups.values()))
        assert len(paths) == 2, paths

        print(f"Scan OK - {len(groups)} group, {len(paths)} duplicate files")
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    main()
