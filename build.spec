# PyInstaller spec for CMC Report Automation - one windowed executable.
# Build (Windows):  build.bat   or  pyinstaller build.spec --noconfirm --clean
# Build (Linux):    build.sh    or  pyinstaller build.spec --noconfirm --clean
# The app uses LibreOffice (installed on the user's PC) for PDFs.
# On Windows, tesseract/ (tesseract.exe, DLLs, tessdata/eng.traineddata) is
# bundled for the "From image..." OCR button.  On Linux, system Tesseract is used.

import os, sys as _sys

_datas = [("app/templates", "app/templates"), ("app/data", "app/data")]
if _sys.platform == "win32" and os.path.isdir("tesseract"):
    _datas.append(("tesseract", "tesseract"))

a = Analysis(
    ["app/main.py"],
    pathex=["."],
    datas=_datas,
    hiddenimports=["docx", "openpyxl", "pypdfium2", "PIL", "app.config.sheet_config"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="CMCReportAutomation",
    console=False,
    upx=False,
)
