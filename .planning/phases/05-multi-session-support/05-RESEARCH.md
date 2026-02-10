# Phase 5: Multi-Session Support - Research

**Researched:** 2026-02-08
**Domain:** MCP HTTP transport, multi-connection WebSocket management, Office.js document identity
**Confidence:** HIGH

## Summary

This phase replaces the stdio MCP transport with Streamable HTTP transport, enabling multiple Claude Code sessions to connect to a single long-running bridge server. The bridge server already exists as an HTTPS+WSS server on port 8443; the MCP endpoint is added as a new HTTP route (`/mcp`) on the same server.

The standard approach is well-established: the `@modelcontextprotocol/sdk` v1.26.0 (already installed) includes `StreamableHTTPServerTransport` which accepts Node.js `IncomingMessage`/`ServerResponse` directly. The SDK provides session management built-in. Each Claude Code session gets its own MCP session with a unique ID, and each session creates a fresh `McpServer` + `StreamableHTTPServerTransport` pair.

The main challenge is TLS certificate trust: Claude Code's Node.js runtime does not automatically trust mkcert's local CA. This requires `NODE_EXTRA_CA_CERTS` pointing to the mkcert root CA certificate.

**Primary recommendation:** Use stateful `StreamableHTTPServerTransport` with per-session McpServer instances, route `/mcp` on the existing HTTPS server, and use auto-detect with per-call override for presentation targeting.

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| `@modelcontextprotocol/sdk` | 1.26.0 (installed) | MCP server + StreamableHTTP transport | Official SDK, already in project |
| `ws` | 8.19.0 (installed) | WebSocket server for add-in connections | Already in project |
| `zod` | 4.3.6 (installed) | Schema validation for MCP tool params | Already in project |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| `@hono/node-server` | 1.19.9 (SDK dep) | Node.js HTTP-to-WebStandard conversion | Automatically used by StreamableHTTPServerTransport |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| StreamableHTTPServerTransport | Plain HTTP + custom JSON-RPC | Massive effort, error-prone, no session management |
| Stateful sessions | Stateless mode | Stateless creates new transport per request; stateful tracks sessions for SSE/notifications |
| Auto-detect targeting | Session-level binding | Binding requires extra state management; auto-detect is zero-friction for single-presentation case |

**Installation:**
```bash
# No new packages needed - everything is already installed
```

## Architecture Patterns

### Recommended Project Structure

No new files needed. All changes are in `server/index.ts` and `addin/app.js`.

```
server/
└── index.ts          # All changes: HTTP routing, multi-connection tracking, MCP session management
addin/
└── app.js            # Add presentation identity reporting to ready message
```

### Pattern 1: Per-Session McpServer + Transport Pairs

**What:** Each MCP HTTP session creates a dedicated `McpServer` and `StreamableHTTPServerTransport` pair. Tools are registered on each server via a factory function.

**When to use:** Always (this is the SDK's recommended pattern for stateful multi-session servers).

**Example:**
```typescript
// Source: modelcontextprotocol/typescript-sdk examples + SDK type definitions
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { randomUUID } from 'node:crypto';

// Session storage
const mcpTransports = new Map<string, StreamableHTTPServerTransport>();

// Factory: creates a new McpServer with all tools registered
function createMcpServer(): McpServer {
  var server = new McpServer({ name: "powerpoint-bridge", version: "0.2.0" });
  // Register tools... (list_presentations, get_presentation, get_slide, execute_officejs)
  return server;
}

// In the HTTPS request handler, for POST /mcp:
async function handleMcpPost(req: IncomingMessage, res: ServerResponse) {
  var body = await parseJsonBody(req);
  var sessionId = req.headers['mcp-session-id'] as string | undefined;

  if (sessionId && mcpTransports.has(sessionId)) {
    // Existing session - route to its transport
    await mcpTransports.get(sessionId)!.handleRequest(req, res, body);
  } else if (!sessionId && isInitializeRequest(body)) {
    // New session - create transport + server
    var transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (sid) => { mcpTransports.set(sid, transport); },
    });
    transport.onclose = () => {
      if (transport.sessionId) mcpTransports.delete(transport.sessionId);
    };
    var server = createMcpServer();
    await server.connect(transport);
    await transport.handleRequest(req, res, body);
  } else {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Bad request: no valid session' }));
  }
}
```

### Pattern 2: Multi-Connection WebSocket Pool

**What:** Replace the single `addinClient` variable with a Map tracking all connected add-in instances, keyed by presentation identifier.

**When to use:** Always (replacing current single-connection model).

**Example:**
```typescript
interface AddinConnection {
  ws: WebSocket;
  ready: boolean;
  presentationId: string;   // file path (stable) or generated ID
  filePath: string | null;  // null if unsaved
}

const addinConnections = new Map<string, AddinConnection>();
let untitledCounter = 0;

// On WebSocket 'ready' message from add-in:
function handleAddinReady(ws: WebSocket, data: { documentUrl?: string }) {
  var filePath = data.documentUrl || null;
  var presentationId = filePath || ('untitled-' + (++untitledCounter));
  var conn: AddinConnection = { ws, ready: true, presentationId, filePath };
  addinConnections.set(presentationId, conn);
}
```

### Pattern 3: Auto-Detect Targeting with Per-Call Override

**What:** If one add-in is connected, auto-target it. If multiple are connected, require explicit `presentationId` parameter. The `list_presentations` tool always works.

**When to use:** Always (recommended for Claude Code UX).

**Rationale:** This is the most natural UX for Claude Code:
- Single presentation (most common): zero friction, identical to current behavior
- Multiple presentations: explicit and discoverable
- No hidden session state to track or lose
- Claude can naturally: list, pick, specify

**Example:**
```typescript
function resolveTarget(presentationId?: string): AddinConnection {
  if (addinConnections.size === 0) {
    throw new Error('No presentations connected. Open a PowerPoint file with the bridge add-in loaded.');
  }
  if (presentationId) {
    var conn = addinConnections.get(presentationId);
    if (!conn) throw new Error('Presentation not found: ' + presentationId + '. Use list_presentations to see connected presentations.');
    return conn;
  }
  if (addinConnections.size === 1) {
    return addinConnections.values().next().value!;
  }
  var ids = [...addinConnections.keys()];
  throw new Error('Multiple presentations connected. Specify presentationId parameter. Available: ' + ids.join(', '));
}
```

### Pattern 4: HTTP Request Routing on Existing HTTPS Server

**What:** Add method-based routing in the existing HTTPS request handler to direct `/mcp` requests to MCP transport while all other requests go to static file serving.

**When to use:** Always (sharing port 8443 between static files, WSS, and MCP HTTP).

**Example:**
```typescript
// Replace serveStatic as the sole request handler:
function handleRequest(req: IncomingMessage, res: ServerResponse): void {
  var url = (req.url ?? '/').split('?')[0];

  if (url === '/mcp') {
    if (req.method === 'POST') handleMcpPost(req, res);
    else if (req.method === 'GET') handleMcpGet(req, res);
    else if (req.method === 'DELETE') handleMcpDelete(req, res);
    else { res.writeHead(405); res.end(); }
    return;
  }

  // Everything else: static files (existing behavior)
  serveStatic(req, res);
}
```

### Pattern 5: JSON Body Parsing for Raw Node.js HTTP

**What:** Parse the JSON body from the raw request stream before routing, then pass it as `parsedBody` to the transport. This is necessary because we need to inspect the body (to check `isInitializeRequest`) before the transport consumes the stream.

**When to use:** For POST requests to `/mcp` only.

**Example:**
```typescript
function parseJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    var chunks: Buffer[] = [];
    req.on('data', (chunk: Buffer) => chunks.push(chunk));
    req.on('end', () => {
      try {
        var body = JSON.parse(Buffer.concat(chunks).toString());
        resolve(body);
      } catch (e) {
        reject(new Error('Invalid JSON body'));
      }
    });
    req.on('error', reject);
  });
}
```

### Anti-Patterns to Avoid

- **Sharing one McpServer across sessions:** The SDK expects one McpServer per transport. Sharing would break session isolation and cause message routing errors.
- **Using stateless mode:** Stateless creates a new transport per request with no session tracking. We need stateful mode so SSE streams and session-scoped behavior work correctly.
- **Pre-parsing body with Express body-parser:** Don't add Express. The raw body parser above is trivial and avoids a massive dependency.
- **Storing transport reference before onsessioninitialized:** The transport's sessionId is only set after the SDK processes the initialize response. Store in the `onsessioninitialized` callback, not at creation time.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| MCP HTTP transport | Custom JSON-RPC over HTTP | `StreamableHTTPServerTransport` | Session management, SSE streaming, protocol compliance all built in |
| Session ID generation | Manual UUID + tracking | `sessionIdGenerator` option | SDK handles session lifecycle, validation, and cleanup |
| Request body parsing | Express middleware | Simple `Buffer.concat` + `JSON.parse` | Only needed for POST /mcp; 10 lines vs adding Express |
| HTTP request routing | Express/Koa framework | Simple `if/else` on `req.url` + `req.method` | Only one route (`/mcp`) with 3 methods; framework is overkill |

**Key insight:** The SDK v1.26.0 handles the hard parts (session management, SSE streaming, protocol compliance). The only custom code needed is: HTTP routing (trivial), body parsing (trivial), and connection pool management (the real logic).

## Common Pitfalls

### Pitfall 1: TLS Certificate Trust

**What goes wrong:** Claude Code's MCP HTTP client (Node.js) does not trust mkcert's local CA by default, causing `SELF_SIGNED_CERT_IN_CHAIN` or `fetch failed` errors when connecting to `https://localhost:8443/mcp`.

**Why it happens:** Node.js uses its bundled CA certificates, not the macOS system Keychain. mkcert installs its CA in the system trust store (trusted by browsers) but Node.js ignores it.

**How to avoid:** Set `NODE_EXTRA_CA_CERTS` to point to mkcert's root CA certificate before launching Claude Code:

```bash
# Find mkcert CA root
mkcert -CAROOT
# Output: $(mkcert -CAROOT)

# Option 1: Environment variable when launching Claude Code
NODE_EXTRA_CA_CERTS="$(mkcert -CAROOT)/rootCA.pem" claude

# Option 2: Add to Claude Code settings.json (persistent)
# In ~/.claude/settings.json:
{
  "env": {
    "NODE_EXTRA_CA_CERTS": "$(mkcert -CAROOT)/rootCA.pem"
  }
}
```

**Warning signs:** MCP connection errors mentioning "fetch failed", "SELF_SIGNED_CERT_IN_CHAIN", or "unable to verify the first certificate".

**Confidence:** HIGH - verified via Claude Code docs and GitHub issue #2899.

### Pitfall 2: Body Stream Consumed Before Transport

**What goes wrong:** If the request body stream is consumed (e.g., by middleware or manual reading) without passing the parsed result to `handleRequest`, the transport gets an empty body and fails to process the request.

**Why it happens:** Node.js request streams can only be read once. The transport's `@hono/node-server` adapter converts `IncomingMessage` to `Request`, which reads the body stream internally. If already consumed, it gets nothing.

**How to avoid:** Always pass the pre-parsed body as the third argument to `handleRequest(req, res, parsedBody)` when you've already read the body to check `isInitializeRequest`.

**Warning signs:** Transport silently fails or returns 400 errors on valid requests.

### Pitfall 3: Storing Transport Before Session Initialization

**What goes wrong:** Storing the transport in the sessions map immediately after creation (before `onsessioninitialized` fires) means the sessionId key is `undefined`.

**Why it happens:** The SDK generates the session ID during the initialize handshake, not at transport construction time. `transport.sessionId` is undefined until the handshake completes.

**How to avoid:** Only store the transport in the `onsessioninitialized` callback:
```typescript
var transport = new StreamableHTTPServerTransport({
  sessionIdGenerator: () => randomUUID(),
  onsessioninitialized: (sessionId) => {
    mcpTransports.set(sessionId, transport);  // Correct: sessionId is now set
  },
});
```

**Warning signs:** Map has `undefined` keys; subsequent requests with valid session IDs get 400 errors.

### Pitfall 4: Unsaved Presentation Identity

**What goes wrong:** `Office.context.document.url` returns null or empty string for unsaved presentations, leaving the add-in unidentifiable.

**Why it happens:** Office.js only provides a URL/path for documents that have been saved to disk at least once.

**How to avoid:** Generate a fallback ID (`untitled-1`, `untitled-2`, etc.) on the server side when the add-in reports no document URL. Update the identifier when the user saves and the add-in reports a file path.

**Warning signs:** Multiple unsaved presentations get the same empty-string key; connections overwrite each other in the map.

### Pitfall 5: Server Must Be Running Before Claude Code

**What goes wrong:** With stdio transport, the MCP server starts automatically when Claude Code launches. With HTTP transport, the bridge server must already be running.

**Why it happens:** Stdio transport launches the server as a subprocess. HTTP transport connects to an existing URL.

**How to avoid:** Clear error messaging when the bridge server is not running. Document the startup procedure. The `.mcp.json` config no longer auto-starts anything.

**Warning signs:** Claude Code shows "MCP server powerpoint-bridge: HTTP Connection error" on startup.

### Pitfall 6: pendingRequests Map Shared Across Sessions

**What goes wrong:** The `pendingRequests` map (tracking WebSocket command responses) is currently global. With multiple sessions sending commands to the same or different add-ins, request IDs must be globally unique.

**Why it happens:** Multiple MCP sessions may issue commands simultaneously. If UUIDs are used (current behavior), this is fine. But if sequential IDs were used, collisions would occur.

**How to avoid:** Keep using `randomUUID()` for command IDs (already implemented). The global pendingRequests map is fine because UUID keys are globally unique regardless of which session originated the command.

**Warning signs:** None with current UUID approach - this is a "don't change what works" note.

## Code Examples

### Complete handleRequest Signature (from SDK v1.26.0)

```typescript
// Source: @modelcontextprotocol/sdk v1.26.0 dist/cjs/server/streamableHttp.d.ts
import { IncomingMessage, ServerResponse } from 'node:http';

handleRequest(
  req: IncomingMessage & { auth?: AuthInfo },
  res: ServerResponse,
  parsedBody?: unknown
): Promise<void>;
```

### StreamableHTTPServerTransport Constructor Options (from SDK v1.26.0)

```typescript
// Source: @modelcontextprotocol/sdk v1.26.0 dist/cjs/server/webStandardStreamableHttp.d.ts
interface StreamableHTTPServerTransportOptions {
  sessionIdGenerator?: () => string;         // Required for stateful mode
  onsessioninitialized?: (sessionId: string) => void | Promise<void>;
  onsessionclosed?: (sessionId: string) => void | Promise<void>;
  enableJsonResponse?: boolean;              // Default false (SSE preferred)
  eventStore?: EventStore;                   // For resumability (optional)
  allowedHosts?: string[];                   // Deprecated
  allowedOrigins?: string[];                 // Deprecated
  enableDnsRebindingProtection?: boolean;    // Deprecated
  retryInterval?: number;                    // SSE retry interval in ms
}
```

### Claude Code .mcp.json for HTTP Transport

```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "type": "http",
      "url": "https://localhost:8443/mcp"
    }
  }
}
```

### Claude Code CLI Command to Add HTTP Server

```bash
claude mcp add --transport http powerpoint-bridge https://localhost:8443/mcp
```

### Add-in Ready Message with Presentation Identity

```javascript
// Source: Office.js API - Office.Document.url property
// In addin/app.js, modify the ws.onopen handler:
ws.onopen = function() {
  reconnectAttempt = 0;
  updateStatus('connected');
  var documentUrl = null;
  try {
    documentUrl = Office.context.document.url || null;
  } catch (e) {
    // May not be available in standalone mode
  }
  ws.send(JSON.stringify({
    type: 'ready',
    documentUrl: documentUrl
  }));
};
```

### Concurrent Access Warning Tracking

```typescript
// Track which sessions target which presentations (for warning)
const sessionPresentationWarnings = new Map<string, Set<string>>();
// mcpSessionId -> Set of presentationIds already warned about

function checkConcurrentWarning(mcpSessionId: string, presentationId: string): string | null {
  // Count other sessions targeting this presentation
  var otherSessions = 0;
  for (var [sid, transport] of mcpTransports) {
    if (sid !== mcpSessionId) {
      // Check if this session has targeted this presentation
      // (tracked separately)
      otherSessions++;
    }
  }
  if (otherSessions === 0) return null;

  var warned = sessionPresentationWarnings.get(mcpSessionId);
  if (warned?.has(presentationId)) return null; // Already warned

  if (!warned) {
    warned = new Set();
    sessionPresentationWarnings.set(mcpSessionId, warned);
  }
  warned.add(presentationId);

  return 'Note: Another session is also connected to this presentation. Changes from either session will apply immediately (last-write-wins).';
}
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| HTTP+SSE transport (2024-11-05 spec) | Streamable HTTP (2025-03-26 spec) | March 2025 | SSE deprecated; Streamable HTTP is the standard |
| Stdio-only MCP | HTTP + stdio transports | March 2025 | Enables remote and multi-session MCP servers |
| `@modelcontextprotocol/sdk` v1.x | v2 pre-alpha (unreleased) | Expected Q1 2026 | v1.26.0 is stable and recommended; v2 not yet released |

**Deprecated/outdated:**
- **HTTP+SSE transport** (protocol version 2024-11-05): Replaced by Streamable HTTP in 2025-03-26 spec
- **`enableDnsRebindingProtection` option**: Deprecated in SDK; use external middleware instead
- **StdioServerTransport for multi-session**: Stdio is 1:1 with the Claude Code process; cannot support multiple sessions

## Discretion Recommendations

### Targeting Mechanism: Auto-Detect with Per-Call Override

**Recommendation:** Add an optional `presentationId` parameter to `get_presentation`, `get_slide`, and `execute_officejs`. If omitted and exactly one add-in is connected, auto-target it. If omitted and multiple are connected, return an error listing available presentations.

**Rationale:**
1. Zero-friction for single-presentation case (most common, identical to current UX)
2. Explicit and discoverable for multi-presentation case
3. No hidden session state to manage or accidentally lose
4. Natural Claude Code workflow: `list_presentations` -> pick one -> pass `presentationId`

**Rejected alternatives:**
- Session-level binding: Adds state management complexity, risk of stale bindings
- Pure per-call (always required): Adds friction to the common single-presentation case

### Generated ID Format for Unsaved Presentations

**Recommendation:** `untitled-{N}` where N is a server-lifetime auto-incrementing counter starting at 1.

**Rationale:** Simple, readable, predictable. The ID is temporary until the presentation is saved, at which point it updates to the file path.

### MCP HTTP Endpoint Mounting

**Recommendation:** Simple `if/else` routing in the existing HTTPS request handler. Check `req.url === '/mcp'` first, then route by HTTP method (POST/GET/DELETE). Everything else falls through to `serveStatic`.

**Rationale:** Only one route with three methods. Adding Express or a router framework would be overkill and add a huge dependency tree.

### Warning Message Content

**Recommendation:** Return a one-time note in the tool response text (not an error):

> Note: Another session is also connected to this presentation. Changes from either session will apply immediately (last-write-wins).

Track per (MCP session, presentation) pair so it only appears once per session per presentation.

## Open Questions

1. **Presentation identity update on save**
   - What we know: `Office.context.document.url` is null for unsaved files, populated for saved files
   - What's unclear: Whether Office.js fires an event when a file is saved (there's no `DocumentSaved` event type). The URL would need to be re-checked.
   - Recommendation: Poll `document.url` on each command execution. If it changes from null to a path, update the connection map. Alternatively, include `documentUrl` in every WebSocket response from the add-in so the server can detect changes passively.

2. **Add-in reconnection identity**
   - What we know: If the WebSocket disconnects and reconnects, the add-in sends a new `ready` message
   - What's unclear: Whether the same presentation gets a new generated ID if it wasn't saved
   - Recommendation: Include the previous `presentationId` in the ready message if available (stored in the add-in's in-memory state). The server can then reassociate the connection.

3. **Multiple windows of same file**
   - What we know: PowerPoint on macOS can open the same file in multiple windows
   - What's unclear: Whether `document.url` would be identical for both
   - Recommendation: Accept that two windows of the same file produce the same presentationId. Commands go to whichever add-in instance registered last (natural last-write-wins behavior).

## Sources

### Primary (HIGH confidence)

- `@modelcontextprotocol/sdk` v1.26.0 installed package - type definitions for `StreamableHTTPServerTransport`, `StreamableHTTPServerTransportOptions`, `isInitializeRequest` read directly from `dist/cjs/server/streamableHttp.d.ts` and `dist/cjs/server/webStandardStreamableHttp.d.ts`
- `@modelcontextprotocol/sdk` v1.26.0 installed package - runtime source read from `dist/esm/server/streamableHttp.js` confirming `@hono/node-server` usage and `handleRequest` implementation
- Import verification: `StreamableHTTPServerTransport` and `isInitializeRequest` successfully imported and confirmed as functions via Node.js runtime test
- MCP Specification (2025-03-26): https://modelcontextprotocol.io/specification/2025-03-26/basic/transports - Streamable HTTP transport spec
- Claude Code MCP docs: https://code.claude.com/docs/en/mcp - `.mcp.json` format, `claude mcp add --transport http` command
- Claude Code network config docs: https://code.claude.com/docs/en/network-config - `NODE_EXTRA_CA_CERTS` for custom CA trust
- Office.js Document.url API: https://learn.microsoft.com/en-us/javascript/api/office/office.document - document.url property

### Secondary (MEDIUM confidence)

- GitHub issue anthropics/claude-code#2899 - Self-signed cert trust issue with local MCP servers, confirming `NODE_EXTRA_CA_CERTS=$(mkcert -CAROOT)/rootCA.pem` as workaround
- TypeScript SDK examples pattern - per-session McpServer + transport creation with `onsessioninitialized` callback (verified against SDK types)

### Tertiary (LOW confidence)

- None - all findings verified against primary or secondary sources

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - all libraries already installed, import paths verified at runtime
- Architecture: HIGH - SDK type definitions and source code read directly, patterns verified against official examples
- Pitfalls: HIGH - TLS trust issue confirmed via Claude Code docs and GitHub issue; body parsing behavior confirmed from SDK source code

**Research date:** 2026-02-08
**Valid until:** 2026-03-08 (30 days - stable SDK, no breaking changes expected)
