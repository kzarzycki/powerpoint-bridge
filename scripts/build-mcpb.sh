#!/bin/bash
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
STAGE_DIR="$REPO_DIR/dist/mcpb-stage"

echo "=== Building PowerPoint Bridge .mcpb ==="

# 1. Build CJS bundle
echo "[build] Building self-contained CJS bundle..."
cd "$REPO_DIR"
npm run build

# 2. Clean staging area
rm -rf "$STAGE_DIR"
mkdir -p "$STAGE_DIR/server" "$STAGE_DIR/addin/assets"

# 3. Copy self-contained CJS bundle (no node_modules needed)
echo "[build] Copying server bundle..."
cp "$REPO_DIR/dist/index.cjs" "$STAGE_DIR/server/index.cjs"

# 4. Copy add-in static files
echo "[build] Copying add-in files..."
cp "$REPO_DIR/addin/index.html" "$REPO_DIR/addin/app.js" "$REPO_DIR/addin/style.css" "$STAGE_DIR/addin/"
cp "$REPO_DIR/addin/manifest.xml" "$REPO_DIR/addin/manifest-https.xml" "$STAGE_DIR/addin/"
cp "$REPO_DIR/addin/assets/"*.png "$STAGE_DIR/addin/assets/"

# 5. Copy manifest and icon
cp "$REPO_DIR/manifest.json" "$STAGE_DIR/"
cp "$REPO_DIR/addin/assets/icon-80.png" "$STAGE_DIR/icon.png"

# 6. Pack as .mcpb
VERSION=$(node -e "process.stdout.write(JSON.parse(require('fs').readFileSync('$REPO_DIR/package.json','utf8')).version)")
MCPB_FILE="$REPO_DIR/powerpoint-bridge-v${VERSION}.mcpb"

echo "[pack] Creating $MCPB_FILE..."
cd "$STAGE_DIR"
npx @anthropic-ai/mcpb pack

# Move the output .mcpb to repo root
PACKED=$(ls "$STAGE_DIR"/*.mcpb 2>/dev/null | head -1)
if [ -n "$PACKED" ]; then
  mv "$PACKED" "$MCPB_FILE"
fi

# Clean up staging
rm -rf "$STAGE_DIR"

echo ""
echo "=== Done ==="
echo "Output: $MCPB_FILE"
echo "Install: Claude Desktop → Settings → Extensions → Advanced → Install Extension"
