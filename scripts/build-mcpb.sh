#!/bin/bash
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
DIST_DIR="$REPO_DIR/dist"

echo "=== Building PowerPoint Bridge .mcpb ==="

# 1. Clean
rm -rf "$DIST_DIR"
mkdir -p "$DIST_DIR/server" "$DIST_DIR/addin/assets"

# 2. Bundle server with esbuild (all TS + deps → single JS file)
echo "[build] Bundling server..."
npx esbuild "$REPO_DIR/server/index.ts" \
  --bundle \
  --platform=node \
  --target=node16 \
  --format=esm \
  --outfile="$DIST_DIR/server/index.js" \
  --external:ws

echo "[build] Server bundle: $(wc -c < "$DIST_DIR/server/index.js" | tr -d ' ') bytes"

# 3. Copy ws dependency (CJS package, must be external for ESM bundle)
echo "[build] Copying ws dependency..."
mkdir -p "$DIST_DIR/node_modules"
cp -r "$REPO_DIR/node_modules/ws" "$DIST_DIR/node_modules/ws"

# 5. Copy add-in static files
echo "[build] Copying add-in files..."
cp "$REPO_DIR/addin/index.html" "$REPO_DIR/addin/app.js" "$REPO_DIR/addin/style.css" "$DIST_DIR/addin/"
cp "$REPO_DIR/addin/manifest.xml" "$REPO_DIR/addin/manifest-https.xml" "$DIST_DIR/addin/"
cp "$REPO_DIR/addin/assets/"*.png "$DIST_DIR/addin/assets/"

# 6. Copy manifest and icon
cp "$REPO_DIR/manifest.json" "$DIST_DIR/"
cp "$REPO_DIR/addin/assets/icon-80.png" "$DIST_DIR/icon.png"

# 7. Pack as .mcpb
VERSION=$(node -e "process.stdout.write(JSON.parse(require('fs').readFileSync('$REPO_DIR/package.json','utf8')).version)")
MCPB_FILE="$REPO_DIR/powerpoint-bridge-v${VERSION}.mcpb"

echo "[pack] Creating $MCPB_FILE..."
cd "$DIST_DIR"
npx @anthropic-ai/mcpb pack

# Move the output .mcpb to repo root if mcpb outputs it in dist
PACKED=$(ls "$DIST_DIR"/*.mcpb 2>/dev/null | head -1)
if [ -n "$PACKED" ]; then
  mv "$PACKED" "$MCPB_FILE"
fi

echo ""
echo "=== Done ==="
echo "Output: $MCPB_FILE"
echo "To install: open the .mcpb file with Claude Desktop"
