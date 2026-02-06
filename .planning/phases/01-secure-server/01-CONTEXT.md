# Phase 1: Secure Server - Context

**Gathered:** 2026-02-06
**Status:** Ready for planning

<domain>
## Phase Boundary

A Node.js server runs on localhost with trusted TLS, serves static files over HTTPS, and accepts WebSocket connections over WSS. This is pure infrastructure — no add-in code, no command protocol, no MCP.

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion

User skipped discussion — all implementation decisions are at Claude's discretion. Key references:

- CLAUDE.md specifies: mkcert for TLS, single Node.js process, TypeScript, certs/ directory
- RESEARCH.md details: mkcert setup commands, WKWebView WSS requirement, localhost cert generation
- Port number, directory layout, npm scripts, build setup — all Claude's choice
- Server startup behavior and console output — Claude's choice

</decisions>

<specifics>
## Specific Ideas

No specific requirements — open to standard approaches. CLAUDE.md and RESEARCH.md provide sufficient technical guidance.

</specifics>

<deferred>
## Deferred Ideas

None — discussion skipped for infrastructure phase.

</deferred>

---

*Phase: 01-secure-server*
*Context gathered: 2026-02-06*
