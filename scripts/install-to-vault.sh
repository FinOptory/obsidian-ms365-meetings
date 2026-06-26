#!/bin/bash
PLUGIN_ID="ms365-meetings"
VAULT="${GDOPS_VAULT_PATH:-$HOME/GDOps}"
DEST="$VAULT/.obsidian/plugins/$PLUGIN_ID"
mkdir -p "$DEST"
cp main.js manifest.json styles.css "$DEST/" 2>/dev/null || true
echo "✅ Installiert in $DEST"
