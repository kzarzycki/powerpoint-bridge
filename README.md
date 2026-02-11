# PowerPoint Bridge

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D24.0.0-brightgreen)](https://nodejs.org)

An MCP server that lets AI assistants manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs.

Unlike file-based tools (python-pptx), PowerPoint Bridge works with presentations that are already open — changes appear instantly, and you keep full access to PowerPoint's UI, animations, and formatting.

## Motivation

This project was inspired by the [Claude in PowerPoint](https://support.anthropic.com/en/articles/11360939-using-claude-in-powerpoint) add-in. The first time I tried it, I was amazed — it edits live, open decks via Office.js, and the results are far better than file-based pptx tools. But it only works inside the add-in, which means no access to CLAUDE.md, skills, or any other Claude Code features. PowerPoint Bridge brings those same Office.js capabilities to Claude Code (and any MCP client) so you get live editing with the full power of your coding environment.

## Architecture

```
AI Assistant  <--MCP HTTP-->  Bridge Server (Node.js)  <--WSS-->  PowerPoint Add-in (Office.js)
                                     |                                      |
                               localhost:3001/mcp                    WKWebView sandbox
                               localhost:8443 (HTTPS)               Office.js API 1.1-1.9
                               serves add-in files                  executes commands on
                               WebSocket server                     live presentation
```

Three components in one repo:

- **`addin/`** — Office.js taskpane add-in that loads inside PowerPoint and connects as a WebSocket client
- **`server/`** — Node.js bridge server: HTTPS + WSS + MCP HTTP transport
- **`certs/`** — Local TLS certificates (generated, gitignored)

## Prerequisites

- **macOS** (primary platform)
- **Node.js >= 24** (uses native TypeScript execution)
- **Microsoft PowerPoint for Mac**
- **mkcert** for local TLS certificates

```bash
brew install mkcert node
```

## Quick Start

### 1. Clone and install

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git
cd powerpoint-bridge
npm install
```

### 2. Generate TLS certificates

PowerPoint's WKWebView requires WSS (secure WebSocket), so we need local TLS certs:

```bash
mkcert -install    # One-time: adds mkcert CA to macOS Keychain (requires password)
npm run setup-certs
```

### 3. Sideload the add-in into PowerPoint

```bash
npm run sideload
```

Then restart PowerPoint. The add-in will appear in the **Home** ribbon tab.

### 4. Start the bridge server

```bash
npm start
```

### 5. Open a presentation and load the add-in

Open any PowerPoint file, then click the bridge add-in button in the ribbon. The taskpane should show "Connected".

## MCP Client Configuration

### Claude Code

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "type": "http",
      "url": "http://localhost:3001/mcp"
    }
  }
}
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "type": "http",
      "url": "http://localhost:3001/mcp"
    }
  }
}
```

### VS Code / Cursor

Add to your workspace `.vscode/mcp.json`:

```json
{
  "servers": {
    "powerpoint-bridge": {
      "type": "http",
      "url": "http://localhost:3001/mcp"
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `list_presentations` | Lists all connected presentations with their IDs and status |
| `get_presentation` | Returns slide structure (IDs, shape counts, shape names/types) |
| `get_slide` | Returns detailed shape info for a slide (text, positions, sizes, fills) |
| `execute_officejs` | Runs arbitrary Office.js code inside the live presentation |

When multiple presentations are open, pass `presentationId` (from `list_presentations`) to target a specific one.

## Limitations

- **No images** — Office.js has no direct image insertion API; workaround is Base64 slide import
- **No charts** — Office.js cannot create charts programmatically
- **No animations** — Not exposed in stable APIs
- **Solid fills only** — No gradients, effects, or shadows
- **Points for positioning** — All position/size values are in points (1 point = 1/72 inch)

## Security

PowerPoint Bridge runs entirely on localhost:

- The HTTPS/WSS server binds to `localhost:8443`
- The MCP HTTP server binds to `localhost:3001`
- TLS certificates are self-signed via mkcert and trusted only on your machine
- No data leaves your machine

**`execute_officejs` runs arbitrary code** inside PowerPoint's Office.js runtime. This is by design — it gives the AI full access to the Office.js API. Only use this with MCP clients you trust.

## Troubleshooting

**"TLS certificate files not found"**
Run `npm run setup-certs` to generate certificates. If this is your first time, also run `mkcert -install` first.

**Add-in not appearing in PowerPoint**
1. Run `npm run sideload` and restart PowerPoint
2. Check that the file exists: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml`

**Add-in shows "Disconnected"**
Make sure the bridge server is running (`npm start`). The add-in auto-reconnects with exponential backoff.

**"Certificate not trusted" in browser**
Run `mkcert -install` to add the mkcert CA to your system Keychain. You may need to enter your macOS password.

## Platform Support

| Platform | Status |
|----------|--------|
| macOS | Supported (primary) |
| Windows | Untested — different sideloading path, may not require WSS |
| Linux | Not supported (no PowerPoint for Linux) |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup and guidelines.

## License

[MIT](LICENSE)
