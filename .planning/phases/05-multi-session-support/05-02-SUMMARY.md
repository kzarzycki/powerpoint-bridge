---
phase: 05-multi-session-support
plan: 02
subsystem: api
tags: [mcp-http, multi-session, streamable-http, presentation-targeting]

# Dependency graph
requires:
  - phase: 05-multi-session-support/01
    provides: Multi-connection WebSocket pool with resolveTarget()
provides:
  - MCP HTTP transport on plain HTTP port 3001
  - list_presentations tool for discovering connected presentations
  - Per-call presentationId targeting on all tools
  - Per-session McpServer instances for concurrent Claude Code sessions
affects: []

# Tech tracking
tech-stack:
  added: [StreamableHTTPServerTransport]
  patterns: [per-session-mcp-factory, http-mcp-transport, concurrent-warning-tracking]

key-files:
  created: []
  modified:
    - server/index.ts
    - .mcp.json
    - package.json

key-decisions:
  - "KD-0502-1: Plain HTTP on port 3001 for MCP instead of HTTPS — Claude Code's fetch doesn't trust mkcert CA"
  - "KD-0502-2: Per-session McpServer+Transport pairs via factory function registerTools()"
  - "KD-0502-3: Auto-detect single presentation, require presentationId for multiple"
  - "KD-0502-4: Concurrent access warning shown once per session per presentation"

patterns-established:
  - "createMcpSession() factory: fresh McpServer + registerTools per HTTP session"
  - "Plain HTTP for MCP, HTTPS+WSS for add-in (split ports for TLS trust)"

# Metrics
duration: 7min
completed: 2026-02-08
---

# Phase 5 Plan 2: MCP HTTP Transport & Presentation Targeting Summary

**Replaced stdio MCP transport with Streamable HTTP on dedicated plain HTTP port, added list_presentations tool and per-call presentationId targeting for multi-session support**

## Performance

- **Duration:** 7 min
- **Started:** 2026-02-08T20:50:00Z
- **Completed:** 2026-02-08T21:15:00Z
- **Tasks:** 3 (2 auto + 1 checkpoint)
- **Files modified:** 3

## Accomplishments
- Replaced `StdioServerTransport` with `StreamableHTTPServerTransport` for MCP HTTP sessions
- Added `list_presentations` tool showing all connected presentations with IDs and status
- Added optional `presentationId` parameter to `get_presentation`, `get_slide`, and `execute_officejs`
- Per-session McpServer instances via `createMcpSession()` factory — multiple Claude Code sessions supported
- Concurrent access warning tracked per (session, presentation) pair, shown once
- Moved MCP endpoint to plain HTTP on port 3001 to avoid TLS trust issues with Claude Code's fetch
- Updated `.mcp.json` to HTTP transport at `http://localhost:3001/mcp`
- Configured `NODE_EXTRA_CA_CERTS` in `~/.claude/settings.json` (retained as backup)
- Bumped package version to 0.2.0

## Task Commits

Each task was committed atomically:

1. **Task 1: Replace stdio with HTTP transport and add targeting tools** - `2c29573` (feat)
2. **Task 2: Update configuration for HTTP transport and TLS trust** - `9094b1a` (chore)
3. **Orchestrator fix: Plain HTTP MCP server for TLS trust** - `a71bd0b` (fix)

## Files Created/Modified
- `server/index.ts` - StreamableHTTPServerTransport, registerTools factory, handleMcpPost/Get/Delete, plain HTTP server on port 3001
- `.mcp.json` - HTTP transport pointing to `http://localhost:3001/mcp`
- `package.json` - Version bumped to 0.2.0

## Decisions Made
- **KD-0502-1:** Plain HTTP on port 3001 for MCP endpoint — Claude Code's Node.js fetch ignores `NODE_EXTRA_CA_CERTS` and can't connect to mkcert HTTPS. Split architecture: HTTPS+WSS on 8443 for add-in, plain HTTP on 3001 for MCP.
- **KD-0502-2:** Per-session McpServer+Transport pairs created via `createMcpSession()` factory. Each HTTP session gets isolated tool instances with session-aware concurrent warnings.
- **KD-0502-3:** `resolveTarget(presentationId?)` auto-detects when one presentation connected, requires explicit ID when multiple are connected. Zero friction for common case.
- **KD-0502-4:** Concurrent access warning shown once per MCP session per presentation, tracked via `sessionConcurrentWarnings` Map.

## Deviations from Plan

- **Plain HTTP instead of HTTPS for MCP:** Originally planned MCP on the same HTTPS server (port 8443). Claude Code's MCP HTTP client (undici-based fetch) does not respect `NODE_EXTRA_CA_CERTS`, so added a separate plain HTTP server on port 3001 for MCP only. HTTPS+WSS on 8443 retained for the Office.js add-in.

## Issues Encountered

- `NODE_EXTRA_CA_CERTS` in `~/.claude/settings.json` env does not affect Claude Code's own process, only spawned subprocesses. Even setting it in shell profile before launching claude didn't help — the MCP HTTP client uses undici fetch which has different TLS behavior.

## User Setup Required

- Bridge server must be running before Claude Code connects (`npm start`)
- Server is no longer auto-started by Claude Code (HTTP transport vs stdio)

## Next Phase Readiness
- All v1.0 milestone phases complete
- Multi-session support verified with live presentations
- Ready for milestone audit

---
*Phase: 05-multi-session-support*
*Completed: 2026-02-08*
