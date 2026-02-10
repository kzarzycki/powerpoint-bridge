# PowerPoint Office.js Bridge

## What This Is

An MCP server that lets AI assistants manipulate live, open PowerPoint presentations on macOS via Office.js APIs. An Office.js add-in connects over WebSocket to a Node.js bridge server, which exposes MCP tools so Claude Code can read slide contents and make precise modifications to an open deck — enabling co-development of presentations in real-time. Supports multiple simultaneous presentations and concurrent MCP sessions. This is the first live-editing MCP bridge for PowerPoint on macOS; all existing solutions use python-pptx (file-based, no live editing).

## Core Value

Claude Code can see what's on a slide and make precise, iterative modifications to a live presentation — co-developing slides with the user, not generating one-shot decks.

## Requirements

### Validated

- TLS certificate generation and trust setup for WSS — v1
- Office.js add-in that loads in PowerPoint taskpane and connects to bridge server via WSS — v1
- Node.js bridge server serving add-in files over HTTPS, running WSS server, and exposing MCP tools — v1
- Manifest XML for sideloading into PowerPoint on macOS — v1
- JSON command/response protocol with request IDs for async matching — v1
- MCP tool: `get_presentation` — returns structured JSON of all slides with shape summaries — v1
- MCP tool: `get_slide` — returns detailed info for one slide (shapes, text, positions, sizes, colors) — v1
- MCP tool: `execute_officejs` — sends arbitrary Office.js code to add-in for execution, returns result — v1
- Add-in executes arbitrary Office.js code in PowerPoint.run() context and returns results — v1
- Connection status visible in add-in taskpane — v1
- Multi-presentation support with per-call targeting — v1
- Multi-session support for concurrent Claude Code sessions — v1
- MCP HTTP transport for multi-session compatibility — v1
- README with architecture overview, setup guide, and MCP client configurations — v2
- MIT LICENSE file — v2
- Comprehensive .gitignore for macOS/Node.js/IDE files — v2
- Biome lint + format configuration with consistent code style — v2
- Vitest test framework with meaningful coverage of bridge server and MCP tools — v2
- GitHub Actions CI (lint, typecheck, test) — v2
- CONTRIBUTING.md with development setup and contribution guide — v2
- Clean up hardcoded/user-specific paths in documentation — v2
- Package.json version 0.1.0 for first public release — v2
- Windows compatibility notes in documentation — v2

### Active

(None — planning next milestone)

### Out of Scope

- Image insertion — Office.js has no direct image API; Base64 slide import workaround is complex
- Chart creation — not exposed in Office.js API
- Animations/transitions — not in stable APIs
- Gradient fills, effects, shadows — only solid fills supported
- Slide master/theme editing — not available in API
- npm packaging / npm registry publication — install from git for now
- OAuth/Microsoft Store submission — sideloading only

## Context

- Shipped v1 (Feb 6-9, 2026) with 931 LOC. Shipped v2 (Feb 10, 2026) with 1,281 LOC.
- Tech stack: Node.js 24 (native TS), ws library, MCP SDK, Office.js API 1.1-1.9, Biome, Vitest.
- Architecture: HTTPS+WSS on port 8443 (add-in), plain HTTP on port 3001 (MCP).
- Server split into 3 files: bridge.ts (ConnectionPool), tools.ts (MCP tools), index.ts (entrypoint).
- 22 tests, GitHub Actions CI, MIT license.
- Office.js PowerPoint API requirement sets 1.1-1.9 are stable on macOS 16.19+.
- macOS runs add-ins in Safari WKWebView (WebKit2) which enforces WSS.
- Sideloading path: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`

## Constraints

- **WSS mandatory**: macOS WKWebView won't connect to ws://localhost, must use wss:// with trusted certs
- **Add-in sandbox**: WKWebView cannot host servers or bind ports, can only be a WebSocket client
- **Solid fills only**: Office.js doesn't support gradients, effects, or shadows
- **Points for positioning**: All shape coordinates in points (1/72 inch)
- **No image API**: Office.js has no direct image insertion; Base64 slide import is only workaround
- **MCP uses plain HTTP**: Claude Code's HTTP client ignores NODE_EXTRA_CA_CERTS, so MCP runs on separate plain HTTP port

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| WebSocket over HTTP polling | Real-time bidirectional communication needed | Good |
| Single Node.js process for HTTPS+WSS+MCP | Simplicity over microservices | Good |
| mkcert for TLS | Simpler than OpenSSL, auto-trusts in Keychain | Good |
| TypeScript for server, plain JS for add-in | WKWebView compatibility requires plain JS in add-in | Good |
| Sideloading only (no Store submission) | Development use case, no distribution needed | Good |
| 3-tool architecture: get_presentation + get_slide + execute_officejs | Maximum capability with minimum tools | Good |
| Port 8443 for HTTPS+WSS | Standard alternative HTTPS port, avoids conflicts | Good |
| AsyncFunction constructor for dynamic code execution | Runs arbitrary code with context and PowerPoint in scope | Good |
| Plain HTTP on port 3001 for MCP | Claude Code's fetch ignores NODE_EXTRA_CA_CERTS | Good |
| Per-session McpServer instances | Isolates concurrent Claude Code sessions | Good |
| Biome over ESLint+Prettier | Single tool, Rust-based, faster, simpler config | Good |
| 3-file server split (bridge/tools/index) | Testability without over-engineering | Good |
| Vitest over Jest/node:test | Native ESM+TS, MCP SDK uses it, InMemoryTransport for testing | Good |
| Keep .planning/ in repo | Educational value showing AI-native engineering process | Good |
| Keep git history | Educational value, no sensitive data in code files | Good |

---
*Last updated: 2026-02-10 after v2 Open Source Release milestone*
