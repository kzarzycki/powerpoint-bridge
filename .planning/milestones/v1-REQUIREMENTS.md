# Requirements Archive: v1 MVP

**Archived:** 2026-02-09
**Status:** SHIPPED

This is the archived requirements specification for v1.
For current requirements, see `.planning/REQUIREMENTS.md` (created for next milestone).

---

**Defined:** 2026-02-06
**Core Value:** Claude Code can see what's on a slide and make precise, iterative modifications to a live presentation — co-developing slides with the user.

## v1 Requirements

Requirements for initial release. Each maps to roadmap phases.

### Infrastructure

- [x] **INFRA-01**: TLS certificates generated via mkcert and trusted in macOS Keychain
- [x] **INFRA-02**: Node.js HTTPS server serves add-in static files on localhost
- [x] **INFRA-03**: WSS server accepts WebSocket connections from add-in
- [x] **INFRA-04**: MCP server exposes tools to Claude Code
- [x] **INFRA-05**: JSON command/response protocol with unique request IDs for async matching

### Add-in

- [x] **ADDIN-01**: HTML taskpane loads inside PowerPoint and initializes Office.js
- [x] **ADDIN-02**: Add-in connects to bridge server via WSS on load
- [x] **ADDIN-03**: Taskpane displays connection status (connected/disconnected)
- [x] **ADDIN-04**: Add-in executes arbitrary Office.js code received via WebSocket and returns results
- [x] **ADDIN-05**: Manifest XML configured for sideloading on macOS

### MCP Tools

- [x] **TOOL-01**: `get_presentation` returns structured JSON of all slides with shape summaries (count, order, IDs, shape names/types)
- [x] **TOOL-02**: `get_slide` returns detailed info for one slide — all shapes with text content, positions, sizes, fill colors, font colors
- [x] **TOOL-03**: `execute_officejs` accepts Office.js code string, executes it in PowerPoint.run() context, and returns the result

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| INFRA-01 | Phase 1: Secure Server | Complete |
| INFRA-02 | Phase 1: Secure Server | Complete |
| INFRA-03 | Phase 1: Secure Server | Complete |
| INFRA-04 | Phase 4: MCP Tools | Complete |
| INFRA-05 | Phase 3: Command Execution | Complete |
| ADDIN-01 | Phase 2: PowerPoint Add-in | Complete |
| ADDIN-02 | Phase 2: PowerPoint Add-in | Complete |
| ADDIN-03 | Phase 2: PowerPoint Add-in | Complete |
| ADDIN-04 | Phase 3: Command Execution | Complete |
| ADDIN-05 | Phase 2: PowerPoint Add-in | Complete |
| TOOL-01 | Phase 4: MCP Tools | Complete |
| TOOL-02 | Phase 4: MCP Tools | Complete |
| TOOL-03 | Phase 4: MCP Tools | Complete |

## Milestone Summary

**Shipped:** 13 of 13 v1 requirements
**Adjusted:** INFRA-04 originally specified stdio transport; Phase 5 evolved this to HTTP transport on port 3001 for multi-session support.
**Dropped:** None

---
*Archived: 2026-02-09 as part of v1 milestone completion*
