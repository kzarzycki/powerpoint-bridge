# Requirements: PowerPoint Office.js Bridge

**Defined:** 2026-02-06
**Core Value:** Claude Code can see what's on a slide and make precise, iterative modifications to a live presentation — co-developing slides with the user.

## v1 Requirements

Requirements for initial release. Each maps to roadmap phases.

### Infrastructure

- [ ] **INFRA-01**: TLS certificates generated via mkcert and trusted in macOS Keychain
- [ ] **INFRA-02**: Node.js HTTPS server serves add-in static files on localhost
- [ ] **INFRA-03**: WSS server accepts WebSocket connections from add-in
- [x] **INFRA-04**: MCP server (stdio transport) exposes tools to Claude Code
- [ ] **INFRA-05**: JSON command/response protocol with unique request IDs for async matching

### Add-in

- [ ] **ADDIN-01**: HTML taskpane loads inside PowerPoint and initializes Office.js
- [ ] **ADDIN-02**: Add-in connects to bridge server via WSS on load
- [ ] **ADDIN-03**: Taskpane displays connection status (connected/disconnected)
- [ ] **ADDIN-04**: Add-in executes arbitrary Office.js code received via WebSocket and returns results
- [ ] **ADDIN-05**: Manifest XML configured for sideloading on macOS

### MCP Tools

- [x] **TOOL-01**: `get_presentation` returns structured JSON of all slides with shape summaries (count, order, IDs, shape names/types)
- [x] **TOOL-02**: `get_slide` returns detailed info for one slide — all shapes with text content, positions, sizes, fill colors, font colors
- [x] **TOOL-03**: `execute_officejs` accepts Office.js code string, executes it in PowerPoint.run() context, and returns the result

## v2 Requirements

Deferred to future iteration based on real usage. Not in current roadmap.

### Reliability

- **REL-01**: Auto-reconnect when WebSocket connection drops
- **REL-02**: Command queuing when add-in is temporarily disconnected
- **REL-03**: Heartbeat/ping to detect stale connections

### Additional Convenience Tools

- **TOOL2-01**: `add_slide` convenience tool (shortcut for common operation)
- **TOOL2-02**: `add_shape` convenience tool with typed parameters
- **TOOL2-03**: Batch execution tool for multiple operations in one call

## Out of Scope

| Feature | Reason |
|---------|--------|
| Image insertion | Office.js has no direct image API; Base64 slide import workaround is complex |
| Chart creation | Not exposed in Office.js JavaScript API |
| Animations/transitions | Not available in stable APIs (1.1-1.9) |
| Gradient fills, effects, shadows | Only solid fills supported by Office.js |
| Slide master/theme editing | Not available in Office.js API |
| npm packaging / public release | Personal tool for now |
| Microsoft Store submission | Sideloading sufficient for development use |
| OAuth / authentication | Single-user localhost tool |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

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

**Coverage:**
- v1 requirements: 13 total
- Mapped to phases: 13
- Unmapped: 0

---
*Requirements defined: 2026-02-06*
*Last updated: 2026-02-08 — All v1 requirements Complete*
