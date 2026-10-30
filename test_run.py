"""Smoke test: create the window, process events briefly, and close it."""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from PySide6.QtWidgets import QApplication

from duplicate_finder import APP_VERSION, DuplicateFinderApp

app = QApplication([])
window = DuplicateFinderApp()
window.show()

deadline = time.time() + 2
while time.time() < deadline:
    app.processEvents()
    time.sleep(0.01)

print(f"GUI OK - v{APP_VERSION}")
window.close()
app.quit()
