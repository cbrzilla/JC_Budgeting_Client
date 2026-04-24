#!/usr/bin/env bash
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INSTALL_ROOT="$HOME/Library/Application Support/JC Budgeting Server"
PLIST="$HOME/Library/LaunchAgents/com.cliffflanders.jcbudgeting.server.plist"
mkdir -p "$INSTALL_ROOT"
rm -rf "$INSTALL_ROOT/app"
cp -R "$SCRIPT_DIR/../manual/server/app" "$INSTALL_ROOT/app"
chmod +x "$INSTALL_ROOT/app/JCBudgeting.Server"
mkdir -p "$(dirname "$PLIST")"
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.cliffflanders.jcbudgeting.server</string>
  <key>ProgramArguments</key>
  <array>
    <string>$INSTALL_ROOT/app/JCBudgeting.Server</string>
    <string>--urls</string>
    <string>http://0.0.0.0:5099</string>
  </array>
  <key>WorkingDirectory</key>
  <string>$INSTALL_ROOT/app</string>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
</dict>
</plist>
EOF
launchctl unload "$PLIST" >/dev/null 2>&1 || true
launchctl load "$PLIST"
echo "Installed JC Budgeting Server LaunchAgent."
echo "Server page: http://localhost:5099/server"