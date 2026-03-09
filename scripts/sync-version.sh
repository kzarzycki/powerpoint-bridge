#!/bin/bash
# Syncs version from package.json into manifest.json and .claude-plugin/plugin.json.
# Called automatically by npm version lifecycle (the "version" script in package.json).
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION=$(node -e "process.stdout.write(JSON.parse(require('fs').readFileSync('$REPO_DIR/package.json','utf8')).version)")

echo "[sync-version] Syncing version $VERSION to manifest.json and plugin.json"

# Update manifest.json
node -e "
const fs = require('fs');
const path = '$REPO_DIR/manifest.json';
const data = JSON.parse(fs.readFileSync(path, 'utf8'));
data.version = '$VERSION';
fs.writeFileSync(path, JSON.stringify(data, null, 2) + '\n');
"

# Update .claude-plugin/plugin.json
node -e "
const fs = require('fs');
const path = '$REPO_DIR/.claude-plugin/plugin.json';
const data = JSON.parse(fs.readFileSync(path, 'utf8'));
data.version = '$VERSION';
fs.writeFileSync(path, JSON.stringify(data, null, 2) + '\n');
"

# Stage the updated files so they're included in npm version's auto-commit
git add "$REPO_DIR/manifest.json" "$REPO_DIR/.claude-plugin/plugin.json"

echo "[sync-version] Done: all version files set to $VERSION"
