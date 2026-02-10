---
created: 2026-02-10T21:15
title: Add ports configuration via environment variables
area: server
files:
  - server/index.ts:20-21
  - addin/app.js:34
  - addin/manifest.xml (9 occurrences of port 8443)
---

## Problem

Ports 8443 (HTTPS+WSS) and 3001 (MCP HTTP) are hardcoded in server/index.ts. The add-in's WSS URL (wss://localhost:8443) is hardcoded in addin/app.js:34, and the manifest.xml has 9 hardcoded references to https://localhost:8443.

## Solution

Two-part approach:

### Server side (simple)
- `PORT` and `MCP_PORT` env vars with current values as defaults
- ~2 lines changed in server/index.ts

### Add-in side (the real work)
The connection chain is: manifest.xml → PowerPoint loads index.html → app.js connects WSS.

**Preferred approach:** Add a port input field to the add-in taskpane UI. Before clicking "Connect", user can edit the port number (defaulting to 8443). The add-in delays WSS connection until the user confirms, rather than auto-connecting on load.

This avoids manifest templating entirely — the manifest always points to 8443 for loading the HTML, but the WSS connection target is user-configurable at runtime.

**Manifest still needs 8443 for loading the add-in HTML.** If someone changes the HTTPS port, they'd still need to update the manifest and re-sideload. But the WSS port (the one that matters for the bridge connection) would be configurable in-UI without touching any files.

### What changes
| Component | Change |
|---|---|
| server/index.ts | Read PORT/MCP_PORT from env vars |
| addin/index.html | Add port input field + connect button |
| addin/app.js | Don't auto-connect; wait for user to confirm port and click connect |
| addin/manifest.xml | No change (always loads from 8443) |

### Deferred for now
Not worth the complexity for v0.1.0 — port 8443 rarely conflicts. If someone hits a conflict, manually editing 2 files + manifest is a 2-minute fix.
