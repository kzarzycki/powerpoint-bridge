#!/bin/bash
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"

echo "=== PowerPoint Bridge Setup ==="
echo ""

# 1. TLS Certificates
if [ -f "$REPO_DIR/certs/localhost.pem" ]; then
  echo "[certs] Already exist, skipping"
else
  if ! command -v mkcert &> /dev/null; then
    echo "[certs] ERROR: mkcert not found. Install with: brew install mkcert"
    exit 1
  fi
  echo "[certs] Generating TLS certificates..."
  mkdir -p "$REPO_DIR/certs"
  mkcert -cert-file "$REPO_DIR/certs/localhost.pem" \
         -key-file "$REPO_DIR/certs/localhost-key.pem" \
         localhost 127.0.0.1 ::1
  echo "[certs] Generated"
fi

# 2. Sideload add-in manifest
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF_DIR"
cp "$REPO_DIR/addin/manifest.xml" "$WEF_DIR/"
echo "[add-in] Manifest sideloaded to PowerPoint"

# 3. Install skill globally
mkdir -p ~/.claude/skills
ln -sfn "$REPO_DIR/.claude/skills/powerpoint-live" ~/.claude/skills/powerpoint-live
echo "[skill] Installed globally at ~/.claude/skills/powerpoint-live"

echo ""
echo "=== Setup complete ==="
echo ""
echo "Next steps:"
echo "  1. If first time: run 'mkcert -install' (requires macOS password)"
echo "  2. Restart PowerPoint to load the add-in"
echo "  3. Start the bridge: npm start"
echo "  4. In any project, ask Claude: 'enable powerpoint mcp in this project'"
