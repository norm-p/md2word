"""
Bootstrap for PyInstaller frozen bundle.

When running as a frozen .exe, set PYPANDOC_PANDOC to the bundled pandoc.exe
so pypandoc and our own _ensure_pandoc() find it without needing PATH.

Import this module before anything else in the PyInstaller entry point.
"""

import os
import sys


def patch_pandoc_path() -> None:
    if getattr(sys, "_MEIPASS", None):
        bundled = os.path.join(sys._MEIPASS, "pandoc.exe")
        if os.path.exists(bundled):
            os.environ["PYPANDOC_PANDOC"] = bundled
