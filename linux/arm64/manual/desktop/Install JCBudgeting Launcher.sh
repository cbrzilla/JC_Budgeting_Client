#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_PATH="$SCRIPT_DIR/JCBudgeting"
ICON_SOURCE="$SCRIPT_DIR/jcb-logo.png"
DESKTOP_DIR="$HOME/.local/share/applications"
ICON_DIR="$HOME/.local/share/icons/hicolor/256x256/apps"
DESKTOP_FILE="$DESKTOP_DIR/jcbudgeting.desktop"
ICON_TARGET="$ICON_DIR/jcbudgeting.png"

mkdir -p "$DESKTOP_DIR" "$ICON_DIR"
cp -f "$ICON_SOURCE" "$ICON_TARGET"
chmod +x "$APP_PATH"

cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=JC Budgeting
GenericName=Budget Planner
X-GNOME-FullName=JC Budgeting
Comment=JC Budgeting Desktop
Exec=$APP_PATH
Icon=$ICON_TARGET
Terminal=false
Categories=Office;Finance;
Keywords=budget;budgeting;finance;money;
StartupNotify=true
StartupWMClass=JCBudgeting
EOF

chmod +x "$DESKTOP_FILE"
update-desktop-database "$DESKTOP_DIR" >/dev/null 2>&1 || true
gtk-update-icon-cache "$HOME/.local/share/icons/hicolor" >/dev/null 2>&1 || true

echo "Installed JCBudgeting launcher for this user."
echo "You can now launch JCBudgeting from the app menu with the proper icon."
