#!/bin/bash
# Mouse Colony - macOS installer. Double-click (or: bash install_mac.command).
# Installs into ~/Library/Application Support/MouseColony, creates ~/Applications/Mouse Colony.app
# and puts the workbook in ~/Documents/Mouse colony (only if it is not there yet).
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"
ROOT="$HOME/Library/Application Support/MouseColony"
APP="$ROOT/app"; VENV="$ROOT/venv"
BUNDLE="$HOME/Applications/Mouse Colony.app"
DOCS="$HOME/Documents/Mouse colony"
echo "=== Mouse Colony - installer ==="

# Python 3.8+ that also has Tkinter (python.org builds include it; Homebrew needs python-tk)
PY=""
for c in /Library/Frameworks/Python.framework/Versions/Current/bin/python3 python3.13 python3.12 python3.11 \
         python3.10 python3 "$HOME/anaconda3/bin/python" "$HOME/miniconda3/bin/python" \
         /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$c" >/dev/null 2>&1 && \
     "$c" -c 'import sys, tkinter; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null; then
    PY="$c"; break
  fi
done
if [ -z "$PY" ]; then
  echo "No Python 3.8+ with Tkinter found. Install Python from https://www.python.org/downloads/"
  echo "(it includes Tkinter) and run this again."
  read -n 1 -s -r -p "Press any key to close"; exit 1
fi
echo "Using $("$PY" --version) ($PY)"

mkdir -p "$APP"
cp "$SRC"/mouse_colony_tool.py "$SRC"/mousecolony.png "$SRC"/mousecolony.ico "$SRC"/requirements.txt "$SRC"/README.md "$SRC"/CAGE_CARDS_FORMAT.md "$SRC"/GOOGLE_SETUP.md "$SRC"/GEMINI_CARD_INSTRUCTIONS.txt "$APP/"
[ -x "$VENV/bin/python" ] || "$PY" -m venv "$VENV"
echo "Installing packages (small, under a minute)..."
"$VENV/bin/python" -m pip install --disable-pip-version-check --upgrade pip >/dev/null
"$VENV/bin/python" -m pip install --disable-pip-version-check -r "$APP/requirements.txt"
"$VENV/bin/python" -c "import tkinter, openpyxl, gspread, google_auth_oauthlib, PIL; print('All packages OK')"
"$VENV/bin/python" -m pip install --disable-pip-version-check -q tkinterdnd2 >/dev/null 2>&1 || \
  echo "Drag & drop add-on not installed - use the Import button instead."

if [ -f "$DOCS/Volgyi_mice_colony.xlsx" ]; then
  echo "Workbook already in $DOCS - left unchanged."
elif [ -f "$SRC/Volgyi_mice_colony.xlsx" ]; then
  mkdir -p "$DOCS"; cp "$SRC/Volgyi_mice_colony.xlsx" "$DOCS/"
  echo "Workbook copied to $DOCS"
fi

# minimal .app bundle that runs the program
mkdir -p "$BUNDLE/Contents/MacOS" "$BUNDLE/Contents/Resources"
cat > "$BUNDLE/Contents/MacOS/MouseColony" <<EOF
#!/bin/bash
exec "$VENV/bin/python" "$APP/mouse_colony_tool.py" "\$@" 2>>"\$HOME/mouse_colony_error.log"
EOF
chmod +x "$BUNDLE/Contents/MacOS/MouseColony"
cat > "$BUNDLE/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>CFBundleName</key><string>Mouse Colony</string>
  <key>CFBundleExecutable</key><string>MouseColony</string>
  <key>CFBundleIdentifier</key><string>hu.pte.mousecolony</string>
  <key>CFBundleIconFile</key><string>mousecolony</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>NSHighResolutionCapable</key><true/>
</dict></plist>
EOF
if command -v iconutil >/dev/null 2>&1; then
  T="$(mktemp -d)/mousecolony.iconset"; mkdir -p "$T"
  for s in 16 32 64 128 256 512; do sips -z $s $s "$APP/mousecolony.png" --out "$T/icon_${s}x${s}.png" >/dev/null 2>&1 || true; done
  iconutil -c icns "$T" -o "$BUNDLE/Contents/Resources/mousecolony.icns" 2>/dev/null || true
fi
touch "$BUNDLE"
echo ""
echo "Done. 'Mouse Colony' is in ~/Applications (also findable with Spotlight)."
echo "To update: unzip the new version and run this again. To remove: delete the app,"
echo "  ~/Library/Application Support/MouseColony and ~/.mouse_colony.json (the workbook stays)."
read -n 1 -s -r -p "Press any key to close"
