#!/usr/bin/env bash
# Build the standalone Linux binary locally.
# Usage:  ./build.sh
# Output: dist/CMCReportAutomation
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

if [ -f ".venv/bin/python" ]; then
    PY=".venv/bin/python"
else
    PY="python3"
fi

echo "Using interpreter: $PY"
"$PY" --version

echo
echo "Installing dependencies (first run can take a minute)..."
"$PY" -m pip install --upgrade pip
"$PY" -m pip install -r requirements.txt -r requirements-build.txt

echo
echo "Installing required system libraries..."
if command -v apt-get >/dev/null 2>&1; then
    sudo apt-get install -y libxcb-cursor0 libegl1
fi

echo
echo "Building the executable..."
"$PY" -m PyInstaller build.spec --noconfirm --clean

echo
echo "============================================================"
echo " Build complete:  dist/CMCReportAutomation"
echo "============================================================"
