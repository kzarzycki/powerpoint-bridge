# Global Installation Guide

One-time setup to install the PowerPoint bridge, add-in, and skill.

## Prerequisites

- **Node.js 24+** (`brew install node`)
- **mkcert** (`brew install mkcert`)
- **PowerPoint for Mac** 16.96+ (for full API support including screenshots)

## Install

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git ~/powerpoint-bridge
cd ~/powerpoint-bridge
npm install
npm run setup
```

`npm run setup` performs:
1. Generates TLS certificates via `mkcert` (required for WSS in WKWebView)
2. Copies add-in manifest to PowerPoint's sideload directory
3. Symlinks the skill to `~/.claude/skills/powerpoint-live`

## Manual Steps

**First-time mkcert users**: Run `mkcert -install` manually before `npm run setup`. It requires an interactive terminal (macOS password prompt) to add the local CA to Keychain.

**Restart PowerPoint** after setup to load the sideloaded add-in. The add-in appears as a taskpane and auto-connects to the bridge server.

## Daily Use

```bash
cd ~/powerpoint-bridge && npm start
```

The server runs on:
- `https://localhost:8443` — serves add-in files + WSS for add-in connection
- `http://localhost:3001/mcp` — MCP endpoint for Claude Code

## Troubleshooting

**Add-in not appearing in PowerPoint**
- Verify manifest exists: `ls ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml`
- Restart PowerPoint completely (Cmd+Q, reopen)

**WSS connection failing**
- Verify certs exist: `ls ~/powerpoint-bridge/certs/localhost.pem`
- Verify CA trusted: `mkcert -install` (interactive terminal required)
- Check server is running: `curl -k https://localhost:8443`

**"Add-in disconnected" errors**
- The add-in auto-reconnects with exponential backoff (500ms to 30s)
- Check the add-in taskpane in PowerPoint — status should show "Connected"
- If stuck, close and reopen the presentation

**MCP not connecting**
- Verify: `curl http://localhost:3001/mcp -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'`
- Should return JSON with `serverInfo`
