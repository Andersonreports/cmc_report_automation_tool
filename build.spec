# PyInstaller spec for CMC Report Automation - one windowed .exe.
# Build:  pyinstaller build.spec --noconfirm --clean
# The exe uses LibreOffice or Microsoft Word installed on the user's PC for PDFs.

a = Analysis(
    ["app/main.py"],
    pathex=["."],
    datas=[("app/templates", "app/templates")],
    hiddenimports=["docx", "pypdfium2", "PIL", "app.config.sheet_config"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="CMCReportAutomation",
    console=False,
    upx=False,
)
