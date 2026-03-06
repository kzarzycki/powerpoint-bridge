# PowerPoint Bridge

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D24.0.0-brightgreen)](https://nodejs.org)

An MCP server that lets AI assistants manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs.

Unlike file-based tools (python-pptx), PowerPoint Bridge works with presentations that are already open — changes appear instantly, and you keep full access to PowerPoint's UI, animations, and formatting.

## Installation

Run as an MCP server via npx (no install needed):

```bash
npx powerpoint-bridge --stdio --bridge
```

Then configure your MCP client:

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "command": "npx",
      "args": ["-y", "powerpoint-bridge", "--stdio", "--bridge"]
    }
  }
}
```

**Claude Code** (`.mcp.json` in project root):
```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "powerpoint-bridge", "--stdio", "--bridge"]
    }
  }
}
```

**Cursor** (`.cursor/mcp.json`):
```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "command": "npx",
      "args": ["-y", "powerpoint-bridge", "--stdio", "--bridge"]
    }
  }
}
```

**VS Code / GitHub Copilot** (`.vscode/mcp.json`):
```json
{
  "servers": {
    "powerpoint-bridge": {
      "command": "npx",
      "args": ["-y", "powerpoint-bridge", "--stdio", "--bridge"]
    }
  }
}
```

**Windsurf** (`~/.codeium/windsurf/mcp_config.json`):
```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "command": "npx",
      "args": ["-y", "powerpoint-bridge", "--stdio", "--bridge"]
    }
  }
}
```

> **Note:** The PowerPoint add-in must still be sideloaded separately. See [Setup](#setup) for details.

## Motivation

This project was inspired by the [Claude in PowerPoint](https://support.anthropic.com/en/articles/11360939-using-claude-in-powerpoint) add-in. The first time I tried it, I was amazed — it edits live, open decks via Office.js, and the results are far better than file-based pptx tools. But it only works inside the add-in, which means no access to CLAUDE.md, skills, or any other Claude Code features. PowerPoint Bridge brings those same Office.js capabilities to Claude Code (and any MCP client) so you get live editing with the full power of your coding environment.

## Architecture

```
AI Assistant  <--MCP HTTP-->  Bridge Server (Node.js)  <--WS-->  PowerPoint Add-in (Office.js)
                                     |                                     |
                               localhost:3001/mcp                   WKWebView sandbox
                               localhost:8080 (HTTP)                Office.js API 1.1-1.9
                               serves add-in files                  executes commands on
                               WebSocket server                     live presentation
```

Three components in one repo:

- **`addin/`** — Office.js taskpane add-in that loads inside PowerPoint and connects as a WebSocket client
- **`server/`** — Node.js bridge server: HTTP + WS + MCP transport (HTTPS/WSS opt-in via `BRIDGE_TLS=1`)
- **`skills/powerpoint-live/`** — Claude Code skill with tool docs, code patterns, and setup guide. Installed globally by `npm run setup`.
- **`certs/`** — Optional local TLS certificates for HTTPS mode (generated, gitignored)

## Prerequisites

- **macOS** (primary platform)
- **Node.js >= 24** (uses native TypeScript execution)
- **Microsoft PowerPoint for Mac**

```bash
brew install node
```

## Install

### Let Claude do it

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git
```

Then tell Claude: "install powerpoint bridge from `<path>`" — it will handle `npm install`, sideloading, and per-project config.

### Manual install

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git
cd powerpoint-bridge
npm install
npm run setup      # Sideloads add-in, installs Claude Code skill
```

Then restart PowerPoint, open a presentation, and click the bridge add-in in the ribbon. Start the server:

```bash
npm start
```

After setup, the `powerpoint-live` skill is globally available. In any project, ask Claude: "enable powerpoint mcp in this project". See the [setup guide](skills/powerpoint-live/references/setup.md) for per-project configuration details.

## Claude Desktop Extension

Alternatively, build and install as a one-click `.mcpb` extension (from source):

```bash
npm run build:mcpb    # produces powerpoint-bridge-v0.1.0.mcpb
open powerpoint-bridge-v0.1.0.mcpb   # opens Claude Desktop installer
```

The extension auto-starts the bridge and auto-sideloads the add-in. Restart PowerPoint after first install.

**Known limitations:**
- **Chat mode only** — Cowork and Code tabs don't load desktop extensions ([upstream bug](https://github.com/anthropics/claude-code/issues/20377))
- **Single instance** — only one bridge can run on port 8080

## Available Tools

| Tool | Description |
|------|-------------|
| `list_presentations` | Lists all connected presentations with their IDs and status |
| `get_presentation` | Returns slide structure (IDs, shape counts, shape names/types) |
| `get_slide` | Returns detailed shape info for a slide (text, positions, sizes, fills) |
| `get_slide_image` | Captures a visual screenshot of a slide as PNG (requires PowerPoint 16.96+) |
| `get_deck_overview` | Returns thumbnails + text for all/selected slides in one call (efficient full-deck review) |
| `copy_slides` | Copies slides between two open presentations (data stays server-side, never in Claude context) |
| `insert_image` | Inserts an image from a file path, URL, or base64 data onto a slide |
| `get_local_copy` | Returns a local file path for the presentation (passthrough for local, exports cloud files to temp .pptx) |
| `execute_officejs` | Runs arbitrary Office.js code inside the live presentation |

When multiple presentations are open, pass `presentationId` (from `list_presentations`) to target a specific one.

## Limitations

- **Limited image control** — Images inserted via Common API (`insert_image` tool), not shape API; positioning works but no shape-level manipulation after insertion
- **No charts** — Office.js cannot create charts programmatically
- **No animations** — Not exposed in stable APIs
- **Solid fills only** — No gradients, effects, or shadows
- **Points for positioning** — All position/size values are in points (1 point = 1/72 inch)

## Security

PowerPoint Bridge runs entirely on localhost:

- The bridge server binds to `localhost:8080` (HTTP) or `localhost:8443` (HTTPS with `BRIDGE_TLS=1`)
- The MCP HTTP server binds to `localhost:3001`
- No data leaves your machine

**`execute_officejs` runs arbitrary code** inside PowerPoint's Office.js runtime. This is by design — it gives the AI full access to the Office.js API. Only use this with MCP clients you trust.

## Troubleshooting

**Add-in not appearing in PowerPoint**
1. Run `npm run sideload` and restart PowerPoint
2. Check that the file exists: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml`

**Add-in shows "Disconnected"**
Make sure the bridge server is running (`npm start`). You can verify with `curl http://localhost:3001/health`. The add-in auto-reconnects with exponential backoff.

**Using HTTPS mode**
If plain HTTP/WS doesn't work in your environment, switch to HTTPS:
1. `brew install mkcert && mkcert -install` (one-time, requires macOS password)
2. `npm run setup-certs` to generate certificates
3. `npm run sideload:https` and restart PowerPoint
4. `BRIDGE_TLS=1 npm start`

## Platform Support

| Platform | Status |
|----------|--------|
| macOS | Supported (primary) |
| Windows | Untested — different sideloading path |
| Linux | Not supported (no PowerPoint for Linux) |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup and guidelines.

## License

[MIT](LICENSE)
