#!/usr/bin/env bash
set -euo pipefail
PLIST="$HOME/Library/LaunchAgents/com.cliffflanders.jcbudgeting.server.plist"
INSTALL_ROOT="$HOME/Library/Application Support/JC Budgeting Server"
launchctl unload "$PLIST" >/dev/null 2>&1 || true
rm -f "$PLIST"
rm -rf "$INSTALL_ROOT/app"
echo "Removed JC Budgeting Server LaunchAgent/app files. Databases/logs were not intentionally deleted."