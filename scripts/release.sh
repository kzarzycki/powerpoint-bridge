#!/bin/bash
# Cut a release: validate, bump version, push tag → GitHub Actions builds .mcpb.
# Usage: npm run release -- patch|minor|major
#    or: bash scripts/release.sh patch|minor|major
set -e

BUMP="${1:-patch}"

if [ "$BUMP" != "patch" ] && [ "$BUMP" != "minor" ] && [ "$BUMP" != "major" ]; then
  echo "Usage: npm run release -- <patch|minor|major>"
  exit 1
fi

# Ensure clean working tree
if [ -n "$(git status --porcelain)" ]; then
  echo "Error: Working tree is not clean. Commit or stash changes first."
  exit 1
fi

# Ensure on main
BRANCH=$(git branch --show-current)
if [ "$BRANCH" != "main" ]; then
  echo "Warning: You are on branch '$BRANCH', not 'main'. Continue? (y/N)"
  read -r CONFIRM
  if [ "$CONFIRM" != "y" ]; then
    exit 1
  fi
fi

# Run checks before bumping
echo "[release] Running checks..."
npm run check

# Bump version (triggers "version" script which syncs manifest + plugin.json)
echo "[release] Bumping $BUMP version..."
npm version "$BUMP"

# Push commit and tag
echo "[release] Pushing to origin..."
git push origin "$BRANCH" --follow-tags

echo ""
echo "=== Release pushed ==="
echo "GitHub Actions will now build the .mcpb and create the GitHub Release."
echo "Monitor at: https://github.com/kzarzycki/powerpoint-mcp/actions"
