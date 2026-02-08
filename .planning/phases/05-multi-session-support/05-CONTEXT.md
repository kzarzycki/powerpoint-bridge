# Phase 5: Multi-Session Support - Context

**Gathered:** 2026-02-08
**Status:** Ready for planning

<domain>
## Phase Boundary

Enable multiple Claude Code sessions to connect to different open PowerPoint presentations simultaneously, and allow multiple sessions to work with the same presentation. Currently the bridge tracks a single add-in connection and uses stdio MCP (1:1 with Claude Code process).

</domain>

<decisions>
## Implementation Decisions

### Connection architecture
- Switch MCP from stdio transport to HTTP transport on the existing HTTPS server
- One long-running bridge server manages all WebSocket connections from PowerPoint add-ins
- Multiple Claude Code sessions connect directly to the bridge server's MCP HTTP endpoint
- `.mcp.json` changes from `command` (stdio) to `url` (HTTP) — e.g., `https://localhost:8443/mcp`
- Server must be started manually before Claude Code (`npm start`) — no auto-start daemon
- Clear error message when bridge server isn't running: "Bridge server not running. Start with: npm start"

### Presentation targeting
- New `list_presentations` MCP tool shows all connected PowerPoint instances with file paths
- Claude's discretion on targeting mechanism: session-level binding, per-call parameter, or auto-detect with override — pick what's most natural for Claude Code UX

### Presentation identity
- Add-in reports its presentation identity using `document.url` from Office.js (provides file path)
- Unsaved/new presentations get a generated ID (e.g., "untitled-1"), updated to file path once saved
- File path is the stable identifier — slide count is dynamic and not used for identification
- Each PowerPoint instance auto-connects to the bridge server on load (existing behavior, no manual step)

### Concurrent access
- Last-write-wins: multiple sessions can read and write to the same presentation freely
- Warning shown once when a session first targets a presentation that another session is already using
- No write locking — Office.js/PowerPoint handles concurrent writes natively
- No server-side command queuing — commands flow straight through

### Claude's Discretion
- Exact targeting mechanism for MCP tools (session binding vs per-call parameter vs auto-detect)
- Internal connection tracking data structures (replacing single `addinClient` variable)
- How the MCP HTTP endpoint is mounted on the existing HTTPS server
- Generated ID format for unsaved presentations
- Warning message content and format

</decisions>

<specifics>
## Specific Ideas

- Server currently has single `addinClient` and `addinReady` variables — needs to become a collection keyed by presentation identity
- `sendCommand()` needs a target parameter to route to the correct add-in WebSocket connection
- Add-in should send its presentation identity (file path) in the `ready` message so the server can register it
- MCP SDK supports StreamableHTTPServerTransport for HTTP transport

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 05-multi-session-support*
*Context gathered: 2026-02-08*
