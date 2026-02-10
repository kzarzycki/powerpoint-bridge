---
created: 2026-02-10T21:15
title: Add ports configuration via environment variables
area: server
files:
  - server/index.ts:20-21
---

## Problem

Ports 8443 (HTTPS+WSS) and 3001 (MCP HTTP) are hardcoded in server/index.ts. Users running other services on these ports have no way to change them without editing source code. The add-in's WSS URL (wss://localhost:8443) is also hardcoded in addin/app.js:34.

## Solution

Support `PORT` and `MCP_PORT` environment variables with current values as defaults. The add-in WSS URL would need to either be configurable at build time or read from a query parameter / config endpoint. TBD on the add-in side — may not be worth the complexity for localhost-only use.
