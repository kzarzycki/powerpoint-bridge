# Roadmap: PowerPoint Office.js Bridge

## Overview

This roadmap delivers a live PowerPoint editing bridge in four phases, each building on the last. Phase 1 establishes the secure server infrastructure (TLS, HTTPS, WSS). Phase 2 gets the Office.js add-in running inside PowerPoint and connected to the server. Phase 3 wires up the command protocol so arbitrary Office.js code can be sent and executed in the live presentation. Phase 4 adds the MCP server and tools so Claude Code can discover, read, and modify presentations through natural conversation.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3, 4): Planned milestone work
- Decimal phases (e.g., 2.1): Urgent insertions (marked with INSERTED)

- [x] **Phase 1: Secure Server** - TLS certs, HTTPS file serving, and WSS endpoint on localhost
- [x] **Phase 2: PowerPoint Add-in** - Office.js taskpane loads in PowerPoint and connects to bridge server
- [x] **Phase 3: Command Execution** - JSON protocol and arbitrary Office.js code execution over WebSocket
- [ ] **Phase 4: MCP Tools** - Stdio MCP server with get_presentation, get_slide, and execute_officejs tools

## Phase Details

### Phase 1: Secure Server
**Goal**: A Node.js server runs on localhost with trusted TLS, serves static files over HTTPS, and accepts WebSocket connections over WSS
**Depends on**: Nothing (first phase)
**Requirements**: INFRA-01, INFRA-02, INFRA-03
**Success Criteria** (what must be TRUE):
  1. Running `mkcert localhost` generates certificate files and they are trusted by macOS Keychain (no manual trust steps needed)
  2. Visiting `https://localhost:PORT` in Safari shows the add-in page without any certificate warning
  3. A test WebSocket client can connect to `wss://localhost:PORT` and receive a response
**Plans**: 2 plans

Plans:
- [x] 01-01-PLAN.md -- Project foundation: config files, npm dependencies, TLS certificates, placeholder page
- [x] 01-02-PLAN.md -- HTTPS + WSS server with static file serving and browser verification

### Phase 2: PowerPoint Add-in
**Goal**: An Office.js add-in loads inside PowerPoint's taskpane and maintains a live WebSocket connection to the bridge server
**Depends on**: Phase 1
**Requirements**: ADDIN-01, ADDIN-02, ADDIN-03, ADDIN-05
**Success Criteria** (what must be TRUE):
  1. Copying the manifest XML to the wef directory makes the add-in appear in PowerPoint
  2. Clicking the add-in in PowerPoint's ribbon opens an HTML taskpane
  3. The taskpane shows "Connected" when the bridge server is running
  4. The taskpane shows "Disconnected" when the bridge server is stopped
**Plans**: 2 plans

Plans:
- [x] 02-01-PLAN.md -- Add-in HTML, Office.js initialization, manifest XML, and icon assets
- [x] 02-02-PLAN.md -- WSS client with reconnection and live PowerPoint verification

### Phase 3: Command Execution
**Goal**: Commands sent over WebSocket execute as Office.js code inside PowerPoint and return structured results
**Depends on**: Phase 2
**Requirements**: INFRA-05, ADDIN-04
**Success Criteria** (what must be TRUE):
  1. Sending a JSON command via WebSocket triggers Office.js code execution inside PowerPoint
  2. The result (or error) returns as a JSON response with the matching request ID
  3. A test command (e.g., get slide count) returns correct data from the live, open presentation
**Plans**: 1 plan

Plans:
- [x] 03-01-PLAN.md -- JSON command protocol and Office.js execution engine

### Phase 4: MCP Tools
**Goal**: Claude Code discovers and uses MCP tools to read presentation structure, inspect individual slides, and execute arbitrary Office.js modifications
**Depends on**: Phase 3
**Requirements**: INFRA-04, TOOL-01, TOOL-02, TOOL-03
**Success Criteria** (what must be TRUE):
  1. Claude Code lists available MCP tools and sees get_presentation, get_slide, and execute_officejs
  2. Calling `get_presentation` returns JSON with all slides, their IDs, and shape summaries (counts, names, types)
  3. Calling `get_slide` with a slide index returns detailed shape data including text content, positions, sizes, and fill colors
  4. Calling `execute_officejs` with Office.js code modifies the live presentation and the change appears in PowerPoint immediately
  5. A round-trip workflow works: read presentation state, write Office.js to modify it, read again to confirm the change
**Plans**: 1 plan

Plans:
- [ ] 04-01-PLAN.md -- MCP stdio server with get_presentation, get_slide, and execute_officejs tools

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Secure Server | 2/2 | Complete | 2026-02-06 |
| 2. PowerPoint Add-in | 2/2 | Complete | 2026-02-07 |
| 3. Command Execution | 1/1 | Complete | 2026-02-08 |
| 4. MCP Tools | 0/1 | Not started | - |
