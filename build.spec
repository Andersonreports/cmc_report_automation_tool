# PyInstaller spec for CMC Report Automation - one windowed .exe.
# Build:  pyinstaller build.spec --noconfirm --clean
# The exe uses LibreOffice or Microsoft Word installed on the user's PC for PDFs.
# tesseract/ (tesseract.exe, its DLLs, tessdata/eng.traineddata) is bundled for the OCR.

a = Analysis(
    ["app/main.py"],
    pathex=["."],
    datas=[("app/templates", "app/templates"), ("app/data", "app/data"),
           ("tesseract", "tesseract")],      # OCR for the Genomic Position "From image..."
    hiddenimports=["docx", "openpyxl", "pypdfium2", "PIL", "app.config.sheet_config"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="CMCReportAutomation",
    console=False,
    upx=False,
)
