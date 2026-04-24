#!/usr/bin/env bash
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_APP="$SCRIPT_DIR/../manual/desktop/JCBudgeting.app"
DEST_DIR="$HOME/Applications"
mkdir -p "$DEST_DIR"
rm -rf "$DEST_DIR/JCBudgeting.app"
cp -R "$SOURCE_APP" "$DEST_DIR/JCBudgeting.app"
echo "Installed JC Budgeting to $DEST_DIR/JCBudgeting.app"