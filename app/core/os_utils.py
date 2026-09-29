"""Small platform helpers: bundled-resource paths and PDF converter lookup."""
from __future__ import annotations

import os
import shutil
import sys


def resource_path(*parts: str) -> str:
    """Path to a bundled read-only file (works from source and from the exe)."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
    return os.path.join(base, *parts)


def find_soffice() -> str | None:
    """Locate LibreOffice's soffice executable, or None."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            path = os.path.join(base, "LibreOffice", "program", "soffice.exe")
            if os.path.exists(path):
                return path
    return None
