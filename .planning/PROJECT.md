# PowerPoint Office.js Bridge

## What This Is

A system that lets Claude Code manipulate live, open PowerPoint presentations on macOS via Office.js APIs. An Office.js add-in connects over WebSocket to a Node.js bridge server, which exposes MCP tools so Claude Code can read slide contents and make precise modifications to an open deck — enabling co-development of presentations in real-time. This is the first such solution; all existing macOS tools use python-pptx (file-based, no live editing).

## Core Value

Claude Code can see what's on a slide and make precise, iterative modifications to a live presentation — co-developing slides with the user, not generating one-shot decks.

## Requirements

### Validated

(None yet — ship to validate)

### Active

- [ ] TLS certificate generation and trust setup for WSS
- [ ] Office.js add-in that loads in PowerPoint taskpane and connects to bridge server via WSS
- [ ] Node.js bridge server serving add-in files over HTTPS, running WSS server, and exposing MCP tools via stdio
- [ ] Manifest XML for sideloading into PowerPoint on macOS
- [ ] JSON command/response protocol with request IDs for async matching
- [ ] MCP tools: create presentations, add/delete slides
- [ ] MCP tools: add geometric shapes, lines, text boxes with position/size
- [ ] MCP tools: set text content and font color on shapes
- [ ] MCP tools: set solid fill colors on shapes
- [ ] MCP tools: read slide structure (count, order)
- [ ] MCP tools: read shape list per slide with text content
- [ ] MCP tools: read shape positions, sizes, and formatting details
- [ ] Connection status visible in add-in taskpane
- [ ] Reconnection logic when WebSocket disconnects

### Out of Scope

- Image insertion — Office.js has no direct image API; workaround via Base64 slide import deferred
- Chart creation — not exposed in Office.js API
- Animations/transitions — not in stable APIs
- Gradient fills, effects, shadows — only solid fills supported
- Slide master/theme editing — not available in API
- Tables — deferred to v2 iteration
- Shape grouping/ungrouping — deferred to v2 iteration
- npm packaging / public release — personal tool for now, may share later
- OAuth/Microsoft Store submission — sideloading only

## Context

- All existing macOS MCP servers for PowerPoint use python-pptx (file-based). Only Windows COM solutions have live editing. This would be the first live-editing MCP bridge for macOS.
- Office.js PowerPoint API requirement sets 1.1-1.9 are stable on macOS 16.19+. Set 1.10 is preview.
- macOS runs add-ins in Safari WKWebView (WebKit2) which enforces WSS — plain ws:// won't connect to localhost.
- Add-in is sandboxed in WKWebView — cannot host servers, can only make outbound connections.
- Positioning uses points (1 point = 1/72 inch).
- Sideloading path: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
- mkcert handles TLS cert generation and auto-trusts in macOS Keychain.
- Anthropic's built-in pptx skill uses PptxGenJS for file generation — this project complements it with live editing.

## Constraints

- **WSS mandatory**: macOS WKWebView won't connect to ws://localhost, must use wss:// with trusted certs
- **Add-in sandbox**: WKWebView cannot host servers or bind ports, can only be a WebSocket client
- **Solid fills only**: Office.js doesn't support gradients, effects, or shadows
- **Points for positioning**: All shape coordinates in points (1/72 inch)
- **No image API**: Office.js has no direct image insertion; Base64 slide import is only workaround
- **Single process**: Bridge server combines HTTPS + WSS + MCP stdio in one Node.js process for simplicity

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| WebSocket over HTTP polling | Real-time bidirectional communication needed for co-development; polling adds latency | — Pending |
| Single Node.js process for HTTPS+WSS+MCP | Simplicity over microservices; one process to start/stop | — Pending |
| mkcert for TLS | Simpler than OpenSSL, auto-trusts in Keychain | — Pending |
| TypeScript for add-in and server | Office.js has good TS types, catches API misuse at compile time | — Pending |
| Sideloading only (no Store submission) | Development use case, no distribution needed | — Pending |
| v1 = Bridge + MCP with read/write tools, iterate from usage | Ship working co-development workflow, expand tools based on real needs | — Pending |

---
*Last updated: 2026-02-06 after initialization*
