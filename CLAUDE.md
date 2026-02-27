# PowerPoint Office.js Bridge

## Project Goal

Build a system that lets Claude Code manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs. First live-editing MCP solution for macOS — all others use python-pptx (file-based).

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

## Usage Documentation

For tool reference, code patterns, and usage — see the **powerpoint-live** skill at `.claude/skills/powerpoint-live/`.

## Technical Constraints

- **WSS mandatory** - macOS WKWebView won't connect to `ws://localhost`, must use `wss://`
- **Add-in cannot host servers** - sandboxed in WKWebView, can only make outbound connections
- **No image API** - Office.js has no direct image insertion; workaround is Base64 slide import
- **No charts** - Office.js cannot create charts
- **No animations** - not exposed in stable APIs
- **Solid fills only** - no gradients, effects, or shadows
- **Points for positioning** - 1 point = 1/72 inch

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
