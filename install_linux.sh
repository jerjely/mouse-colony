#!/bin/bash
# Mouse Colony - installer for Linux / Raspberry Pi OS.
#   In a terminal, in the unzipped MouseColony folder:   bash install_linux.sh
# Installs into ~/.local/share/MouseColony, adds a "Mouse Colony" entry to the desktop menu (Office)
# and puts the workbook in ~/Documents/Mouse colony (only if it is not there yet).
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"
ROOT="${XDG_DATA_HOME:-$HOME/.local/share}/MouseColony"
APP="$ROOT/app"; VENV="$ROOT/venv"
DOCS="$HOME/Documents/Mouse colony"
echo "=== Mouse Colony - installer (Linux / Raspberry Pi) ==="

PY="$(command -v python3 || true)"
if [ -z "$PY" ] || ! "$PY" -c 'import tkinter, venv, ensurepip' 2>/dev/null; then
  echo "Python's window toolkit (Tkinter) or venv is missing - installing it (asks for your password)..."
  sudo apt-get update -qq
  sudo apt-get install -y python3 python3-tk python3-venv
  PY="$(command -v python3)"
fi
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' || {
  echo "Python 3.8 or newer is needed."; exit 1; }
echo "Using $("$PY" --version)"

mkdir -p "$APP"
cp "$SRC"/mouse_colony_tool.py "$SRC"/mousecolony.png "$SRC"/requirements.txt "$SRC"/README.md \
   "$SRC"/CAGE_CARDS_FORMAT.md "$SRC"/GOOGLE_SETUP.md "$SRC"/GEMINI_CARD_INSTRUCTIONS.txt "$APP/"
[ -x "$VENV/bin/python" ] || "$PY" -m venv "$VENV"
echo "Installing packages (small, under a minute)..."
"$VENV/bin/python" -m pip install --disable-pip-version-check -q --upgrade pip >/dev/null 2>&1 || true
"$VENV/bin/python" -m pip install --disable-pip-version-check -q -r "$APP/requirements.txt"
"$VENV/bin/python" -c "import tkinter, openpyxl, gspread, google_auth_oauthlib, PIL; print('All packages OK')"
"$VENV/bin/python" -m pip install --disable-pip-version-check -q tkinterdnd2 >/dev/null 2>&1 && \
  "$VENV/bin/python" -c "import tkinterdnd2" >/dev/null 2>&1 || \
  echo "Drag & drop add-on not available here - use the Import button instead."

if [ -f "$DOCS/Volgyi_mice_colony.xlsx" ]; then
  echo "Workbook already in $DOCS - left unchanged."
elif [ -f "$SRC/Volgyi_mice_colony.xlsx" ]; then
  mkdir -p "$DOCS"; cp "$SRC/Volgyi_mice_colony.xlsx" "$DOCS/"
  echo "Workbook copied to $DOCS"
fi

# launcher + desktop menu entry
cat > "$ROOT/mouse-colony" <<EOF
#!/bin/bash
exec "$VENV/bin/python" "$APP/mouse_colony_tool.py" "\$@" 2>>"\$HOME/mouse_colony_error.log"
EOF
chmod +x "$ROOT/mouse-colony"
MENU="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
mkdir -p "$MENU"
cat > "$MENU/mouse-colony.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Mouse Colony
Comment=Breeder ages, relatedness, genotypes and cage history
Exec="$ROOT/mouse-colony"
Icon=$APP/mousecolony.png
Terminal=false
Categories=Office;Science;
EOF
if [ -d "$HOME/Desktop" ]; then
  cp "$MENU/mouse-colony.desktop" "$HOME/Desktop/" && chmod +x "$HOME/Desktop/mouse-colony.desktop"
  command -v gio >/dev/null 2>&1 && gio set "$HOME/Desktop/mouse-colony.desktop" metadata::trusted true 2>/dev/null || true
fi
command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database "$MENU" >/dev/null 2>&1 || true

echo ""
echo "Done. Start 'Mouse Colony' from the menu (Office) or the desktop icon,"
echo "or in a terminal:  $ROOT/mouse-colony"
echo "To update: unzip the new version and run this again."
echo "To remove: rm -rf \"$ROOT\" \"$MENU/mouse-colony.desktop\" ~/Desktop/mouse-colony.desktop ~/.mouse_colony.json"
