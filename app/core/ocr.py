"""Read the NC accession (e.g. NC_000013.11) from a screenshot / photo.

Uses Tesseract OCR when it is installed (or bundled next to the app); otherwise the
text recognition built into Windows 10/11 (through PowerShell). The whole screenshot
can be given, no cropping: every NC accession found in the text is a candidate, and
the one that belongs to the variant's chromosome is preferred. It is only a backup
for the Genomic Position: the caller checks the result against the chromosome.
"""
from __future__ import annotations

import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile

_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($env:OCR_IMAGE)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$result.Text
"""
_ENCODED = base64.b64encode(_SCRIPT.encode("utf-16-le")).decode()
# NC_0000NN.V, tolerant of what OCR does to it: the underscore read as a space / dash,
# the zeros as the letter O, a dot read as a comma.
_NC = re.compile(r"N\s?C\s?[_\-\s.]?\s?([0-9OoIlSB]{6})\s?[.,]\s?(\d{1,2})")
_FIX = str.maketrans("OoIlSB", "001158")
_VALID = set(range(1, 25)) | {12920}            # chromosomes 1-22, X, Y, mitochondrial


class OcrError(Exception):
    pass


def find_tesseract() -> str:
    """Path of tesseract.exe: on PATH, in the usual install folders, or bundled."""
    from .os_utils import resource_path
    found = shutil.which("tesseract")
    if found:
        return found
    roots = [os.environ.get("LOCALAPPDATA", ""), os.environ.get("ProgramFiles", ""),
             os.environ.get("ProgramFiles(x86)", "")]
    paths = [os.path.join(r, *sub) for r in roots if r
             for sub in (("Programs", "Tesseract-OCR", "tesseract.exe"),
                         ("Tesseract-OCR", "tesseract.exe"))]
    paths.append(resource_path("tesseract", "tesseract.exe"))
    return next((p for p in paths if os.path.exists(p)), "")


def _run(cmd: list[str], **kw) -> subprocess.CompletedProcess:
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=90, creationflags=flags, **kw)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise OcrError(f"Text recognition is not available: {e}") from e


def _tesseract(exe: str, path: str, psm: str) -> str:
    done = _run([exe, path, "stdout", "--psm", psm])
    return done.stdout if done.returncode == 0 else ""


def _windows_ocr(path: str) -> str:
    env = dict(os.environ, OCR_IMAGE=os.path.abspath(path))   # WinRT needs backslashes
    done = _run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-EncodedCommand", _ENCODED], env=env)
    if done.returncode != 0:
        raise OcrError("Windows text recognition could not read the image.")
    return done.stdout


def find_ncs(text: str) -> list[str]:
    """Every NC accession in `text` ('NC_000013.11'), in order, without repeats."""
    out = []
    for m in _NC.finditer(text):
        digits = m.group(1).translate(_FIX)
        nc = f"NC_{digits}.{m.group(2)}"
        if int(digits) in _VALID and nc not in out:
            out.append(nc)
    return out


def find_nc(text: str) -> str:
    """The first NC accession in `text`, or ''."""
    return next(iter(find_ncs(text)), "")


def nc_from_image(path: str, expected: int = 0) -> str:
    """The NC accession in the image (a whole screenshot is fine), or '' if none was
    recognised. `expected` is the number of the variant's chromosome accession (13
    for NC_000013.x): an accession with that number is preferred over any other."""
    from PIL import Image
    exe = find_tesseract()
    if not exe and sys.platform != "win32":
        raise OcrError("Install Tesseract OCR to read an image.")
    first = ""
    with Image.open(path) as im:
        im = im.convert("RGB")
        with tempfile.TemporaryDirectory() as tmp:
            for scale in (2, 3, 1.5):           # small text is often missed at 1x
                k = min(scale, 5000 / max(im.size))
                big = im.resize((int(im.width * k), int(im.height * k)), Image.LANCZOS)
                png = os.path.join(tmp, f"ocr{scale}.png")
                big.save(png)
                texts = ([_tesseract(exe, png, psm) for psm in ("11", "6")] if exe
                         else [_windows_ocr(png)])
                for text in texts:
                    for nc in find_ncs(text):
                        if not expected or int(nc[3:9]) == expected:
                            return nc
                        first = first or nc
    return first
