#!/usr/bin/env bash
# CMC Report Automation - Linux installer (Ubuntu / Debian).
#
#   tar -xzf CMCReportAutomation-linux.tar.gz
#   cd CMCReportAutomation-linux
#   ./install.sh
#
# Installs the app for the current user (menu entry + desktop icon), plus the
# system packages it needs: LibreOffice Writer (preview and PDF export), the
# Qt library PySide6 needs, and the report fonts (Arial, Trebuchet MS,
# Calibri-compatible Carlito). The package step asks for your password once.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$HOME/.local/opt/cmc-report-automation"
BIN="$APP_DIR/CMCReportAutomation"
MENU="$HOME/.local/share/applications/cmc-report-automation.desktop"

echo "== Installing CMC Report Automation =="
mkdir -p "$APP_DIR" "$(dirname "$MENU")"
install -m 755 "$HERE/CMCReportAutomation" "$BIN"

cat > "$MENU" <<EOF
[Desktop Entry]
Type=Application
Name=CMC Report Automation
Comment=Whole Exome Sequencing report automation
Exec=$BIN
Icon=x-office-document
Terminal=false
Categories=Office;
EOF
chmod 644 "$MENU"

DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")"
if [ -d "$DESKTOP_DIR" ]; then
  cp "$MENU" "$DESKTOP_DIR/cmc-report-automation.desktop"
  chmod +x "$DESKTOP_DIR/cmc-report-automation.desktop"
  # GNOME: mark the desktop launcher as trusted so it opens on double-click.
  gio set "$DESKTOP_DIR/cmc-report-automation.desktop" metadata::trusted true 2>/dev/null || true
fi
echo "App installed: $BIN"

if command -v apt-get >/dev/null 2>&1; then
  echo "== Installing required system packages (your password may be asked) =="
  sudo apt-get update -y
  sudo apt-get install -y libreoffice-writer libxcb-cursor0 fonts-liberation fonts-crosextra-carlito
  # Real Arial / Trebuchet MS (downloads the free Microsoft core fonts).
  echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" \
    | sudo debconf-set-selections
  if ! sudo apt-get install -y ttf-mscorefonts-installer; then
    echo "NOTE: Microsoft core fonts could not be installed (no internet?)."
    echo "      Reports still work, using the same-size Liberation fonts."
  fi
  fc-cache -f >/dev/null 2>&1 || true
else
  echo "NOTE: not an apt-based system. Install LibreOffice Writer, libxcb-cursor,"
  echo "      and the Liberation / Carlito / Microsoft core fonts with your package manager."
fi

echo
echo "Done. Open 'CMC Report Automation' from the applications menu or the desktop icon."
