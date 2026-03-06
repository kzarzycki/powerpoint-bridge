#!/bin/bash
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"

echo "=== PowerPoint Bridge Setup ==="
echo ""

# 1. Sideload add-in manifest (HTTP by default)
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF_DIR"
cp "$REPO_DIR/addin/manifest.xml" "$WEF_DIR/"
echo "[add-in] Manifest sideloaded to PowerPoint (HTTP mode)"

# 2. Install skill globally
mkdir -p ~/.claude/skills
ln -sfn "$REPO_DIR/skills/powerpoint-live" ~/.claude/skills/powerpoint-live
echo "[skill] Installed globally at ~/.claude/skills/powerpoint-live"

echo ""
echo "=== Setup complete ==="
echo ""
echo "Next steps:"
echo "  1. Restart PowerPoint to load the add-in"
echo "  2. Start the bridge: npm start"
echo "  3. In any project, ask Claude: 'enable powerpoint mcp in this project'"
echo ""
echo "Optional: To use HTTPS/WSS instead of HTTP/WS:"
echo "  brew install mkcert && mkcert -install"
echo "  npm run setup-certs"
echo "  npm run sideload:https"
echo "  BRIDGE_TLS=1 npm start"
