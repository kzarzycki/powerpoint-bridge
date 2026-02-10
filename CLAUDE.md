# PowerPoint Office.js Bridge

## Project Goal

Build a system that lets Claude Code manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs. This is the first such solution - all existing macOS tools use python-pptx (file-based, no live editing).

## Architecture

```
Claude Code  <--MCP HTTP-->  Bridge Server (Node.js)  <--WSS-->  PowerPoint Add-in (Office.js)
                                    |                                      |
                              localhost:3001/mcp                    WKWebView sandbox
                              localhost:8443 (HTTPS)               Office.js API 1.1-1.9
                              serves add-in files                  executes commands on
                              WebSocket server                     live presentation
```

Three components in one repo:

### 1. `addin/` - Office.js PowerPoint Add-in
- HTML/CSS/JS taskpane that loads inside PowerPoint
- Connects as WebSocket **client** to bridge server on load
- Receives JSON commands, executes Office.js API calls, returns results
- Manifest XML for sideloading

### 2. `server/` - Bridge Server (Node.js)
- HTTPS server serving add-in static files
- WSS (secure WebSocket) server for add-in connection
- MCP server (HTTP transport on port 3001) exposing tools to Claude Code
- All three roles in one process for simplicity

### 3. `certs/` - Local TLS Certificates
- Generated via `mkcert` for localhost
- Required because WKWebView enforces WSS (no plain ws://)
- `.gitignore`d

## Technical Constraints

- **WSS mandatory** - macOS WKWebView won't connect to `ws://localhost`, must use `wss://`
- **Add-in cannot host servers** - it's sandboxed in WKWebView, can only make outbound connections
- **No image API** - Office.js has no direct image insertion; workaround is Base64 slide import
- **No charts** - Office.js cannot create charts
- **No animations** - not exposed in stable APIs
- **Solid fills only** - no gradients, effects, or shadows
- **Points for positioning** - 1 point = 1/72 inch

## Available Office.js Capabilities (Requirement Sets 1.1-1.9)

What we CAN do:
- Create presentations, add/delete slides
- Add geometric shapes (rectangle, circle, triangle, etc.), lines, text boxes
- Position and size shapes (left, top, width, height in points)
- Set text content and font color
- Set solid fill colors
- Group/ungroup shapes
- Create and format tables
- Manage hyperlinks
- Set custom properties and metadata tags
- Select slides, shapes, text ranges programmatically
- Insert slides from other presentations (Base64)

## Build Order

### Phase 1: Minimal Viable Bridge
1. Generate TLS certs with `mkcert`
2. Create bare-bones add-in (HTML + Office.js + WS client)
3. Create Node.js server (HTTPS + WSS + static file serving)
4. Create manifest.xml and sideload into PowerPoint
5. Test: add-in connects, send a command, shape appears

### Phase 2: MCP Integration
6. Add MCP server (stdio) to the Node.js bridge server
7. Define initial tool set: `add_slide`, `add_shape`, `set_text`, `get_slides`
8. Configure in Claude Code's MCP settings
9. Test: Claude Code creates shapes via natural language

### Phase 3: Full Tool Set
10. Expand tools: tables, formatting, positioning, slide management
11. Add read operations: get slide contents, shape properties
12. Error handling and reconnection logic
13. Command queuing for when add-in disconnects

## Project Setup

```bash
# Prerequisites
brew install mkcert node

# Generate certs
mkcert -install
mkdir -p certs
cd certs && mkcert localhost 127.0.0.1 ::1 && cd ..

# Install dependencies
npm install

# Start bridge server
npm start

# Sideload add-in (copy manifest to PowerPoint's wef directory)
cp addin/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

## MCP Configuration (for Claude Code)

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

## Key Technical Decisions

1. **Single Node.js process** for HTTPS + WSS + MCP (simplicity over microservices)
2. **TypeScript** for add-in and server (Office.js has good TS types)
3. **JSON command protocol** with request IDs for async response matching
4. **mkcert** for TLS (simpler than OpenSSL, auto-trusts in Keychain)
5. **Sideloading** for development (no Microsoft store submission needed)

## Command Protocol (WebSocket Messages)

```typescript
// Client (add-in) → Server
interface WSMessage {
  type: 'response' | 'error' | 'ready';
  id?: string;        // matches request ID
  data?: any;
}

// Server → Client (add-in)
interface WSCommand {
  type: 'command';
  id: string;         // unique request ID
  action: string;     // e.g. 'addShape', 'setText', 'getSlides'
  params: Record<string, any>;
}
```

## References

See `RESEARCH.md` for full research findings including:
- Detailed API capabilities per requirement set
- Code examples for all shape/text/table operations
- Existing solutions comparison
- macOS-specific issues and workarounds
- All relevant Microsoft documentation and GitHub issue links
