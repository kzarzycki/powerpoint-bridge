# Phase 1: Secure Server - Research

**Researched:** 2026-02-06
**Domain:** Node.js HTTPS + WSS server with mkcert TLS on macOS
**Confidence:** HIGH

## Summary

This phase creates the secure infrastructure layer: TLS certificates via mkcert, an HTTPS server for static file serving, and a WSS (WebSocket Secure) endpoint. All three run in a single Node.js process on localhost.

The standard approach is straightforward: mkcert generates locally-trusted certificates, Node.js's native `node:https` module creates the server, and the `ws` library provides WebSocket server functionality (Node.js 24 has native WebSocket *client* but still no native server). Node.js 24.8.0 supports TypeScript natively via type stripping (stable, no flags needed), which eliminates the need for a build step during development.

Key constraint: macOS 11.3+ requires a password prompt when running `mkcert -install` to trust the CA in Keychain. This is a one-time interactive step, not automatable. The mkcert CA, once installed in the system trust store, is trusted by Safari and by WKWebView (which PowerPoint uses for add-ins on macOS).

**Primary recommendation:** Use Node.js native TypeScript execution (no build step), native `node:https` for the HTTPS server (no Express), and `ws` library for WSS -- all in a single process.

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| Node.js | 24.8.0 (installed) | Runtime with native TS | Already installed; native TypeScript type stripping is stable |
| ws | 8.19.0 | WebSocket server | Only viable option; Node.js 24 has no native WS server |
| mkcert | 1.4.4 (via brew) | TLS certificate generation | Zero-config localhost CA; auto-trusts in macOS Keychain |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| @types/ws | 8.18.1 | TypeScript definitions for ws | Always -- needed for type checking in editor |
| @types/node | 22+ | Node.js TypeScript definitions | Always -- needed for `node:https`, `node:fs` types |
| typescript | 5.9.3 | Type checker (not runtime) | `npx tsc --noEmit` for CI/type checking only; Node.js runs .ts directly |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Native `node:https` | Express | Express adds unnecessary dependency for serving ~3 static files; native is simpler |
| `ws` library | socket.io | socket.io adds protocol overhead, fallback transports not needed; ws is lighter |
| Node.js native TS | tsx / ts-node | tsx adds a dependency; Node 24 runs .ts natively and stably |
| mkcert | OpenSSL manual | OpenSSL requires manual Keychain trust steps; mkcert handles this automatically |

**Installation:**
```bash
# System tools (one-time)
brew install mkcert

# Project dependencies
npm install ws
npm install --save-dev @types/ws @types/node typescript
```

## Architecture Patterns

### Recommended Project Structure
```
powerpoint-bridge/
├── server/
│   └── index.ts         # Single entry point: HTTPS + WSS + static serving
├── addin/
│   └── index.html       # Placeholder static file (proves HTTPS serving works)
├── certs/
│   ├── localhost.pem     # Generated cert (gitignored)
│   └── localhost-key.pem # Generated key (gitignored)
├── package.json
├── tsconfig.json
└── .gitignore
```

### Pattern 1: Single-Process HTTPS + WSS Server

**What:** Create one HTTPS server, attach WebSocket server to it, and serve static files -- all in one `node:https` server instance.
**When to use:** Always for this project. The CLAUDE.md specifies single-process architecture.

```typescript
// Source: ws GitHub README + Node.js https docs
import { createServer } from 'node:https';
import { readFileSync } from 'node:fs';
import { WebSocketServer } from 'ws';

const server = createServer({
  cert: readFileSync('./certs/localhost.pem'),
  key: readFileSync('./certs/localhost-key.pem'),
});

// Attach WSS to the same server
const wss = new WebSocketServer({ server });

// Handle HTTPS requests (static files)
server.on('request', (req, res) => {
  // serve static files from addin/ directory
});

// Handle WebSocket connections
wss.on('connection', (ws) => {
  ws.on('message', (data) => { /* handle */ });
});

server.listen(8443);
```

**Why this pattern:** The `ws` library's `WebSocketServer` accepts a `{ server }` option that shares the HTTP(S) server's port. WebSocket connections arrive as HTTP upgrade requests on the same port. No need for separate ports.

### Pattern 2: Native TypeScript Execution

**What:** Run `.ts` files directly with `node server/index.ts` -- no compilation step.
**When to use:** Always for development. Node.js 24.8.0 has stable type stripping.

**Constraints (verified on this machine):**
- Interfaces, type annotations, type aliases, generics -- all work
- `enum` does NOT work (use `as const` objects or string union types instead)
- File imports MUST include `.ts` extension: `import { foo } from './bar.ts'`
- `tsconfig.json` is IGNORED by Node.js -- it is only for editor/tsc type checking
- `.ts` files follow `package.json` `"type"` field for ESM vs CJS

### Pattern 3: MIME Type Map for Static Serving

**What:** A simple lookup object mapping file extensions to Content-Type headers.
**When to use:** When serving static files without Express.

```typescript
const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

const DEFAULT_MIME = 'application/octet-stream';
```

### Anti-Patterns to Avoid

- **Using `enum` in TypeScript:** Node.js 24 type stripping does not support enums. Use `as const` objects: `const Status = { CONNECTED: 'connected', DISCONNECTED: 'disconnected' } as const;`
- **Omitting `.ts` in imports:** Node.js native TS requires explicit `.ts` extensions in import paths. `import { foo } from './bar'` will fail; use `import { foo } from './bar.ts'`.
- **Using Express for 3 static files:** Adds an unnecessary 57-dependency package for serving index.html, a CSS file, and a JS file. Native `node:https` + `node:fs` handles this in ~30 lines.
- **Separate ports for HTTPS and WSS:** The `ws` library shares the HTTPS server instance. One port handles both.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| TLS certificates | Manual OpenSSL commands + Keychain trust | `mkcert` | mkcert handles CA creation, Keychain trust, and cert generation in 2 commands |
| WebSocket server | Raw HTTP upgrade handling | `ws` library | WebSocket protocol is complex (framing, masking, fragmentation, ping/pong); ws handles it correctly |
| WebSocket protocol | Custom message framing | `ws` library | RFC 6455 compliance requires handling continuation frames, close handshake, etc. |
| MIME type detection | npm package (mime, mime-types) | Simple const object | Only ~6 file types needed for an Office add-in; a lookup object is simpler than a dependency |

**Key insight:** The server itself is simple glue code. The complexity is in TLS trust (solved by mkcert) and WebSocket protocol (solved by ws). Everything else is standard Node.js.

## Common Pitfalls

### Pitfall 1: mkcert -install Requires Password on macOS 11.3+
**What goes wrong:** Running `mkcert -install` silently fails or hangs in CI/automation because macOS requires interactive authentication to modify the system trust store.
**Why it happens:** Apple changed security requirements in Big Sur 11.3 -- modifying certificate trust settings now requires administrator authentication even when running as root.
**How to avoid:** Document this as an interactive one-time setup step. The user must run `mkcert -install` manually and enter their password when prompted. This only needs to happen once per machine.
**Warning signs:** `mkcert -install` outputs an error or the cert is not trusted in Safari.

### Pitfall 2: Certificate File Paths in .gitignore
**What goes wrong:** Certificates get committed to git, or the server crashes because cert files don't exist on a fresh clone.
**Why it happens:** Forgetting to gitignore `certs/` or not handling missing cert files gracefully.
**How to avoid:** Add `certs/` to `.gitignore`. Add a startup check in the server that prints clear instructions if cert files are missing.
**Warning signs:** `git status` shows .pem files; server crashes with ENOENT on fresh clone.

### Pitfall 3: Wrong Certificate Filenames
**What goes wrong:** mkcert's default output filenames depend on the number of SANs provided. Running `mkcert localhost 127.0.0.1 ::1` produces `localhost+2.pem` and `localhost+2-key.pem`, NOT `localhost.pem`.
**Why it happens:** mkcert auto-names files based on first hostname + count of additional names.
**How to avoid:** Use `-cert-file` and `-key-file` flags to control output filenames: `mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1`
**Warning signs:** Server fails to start with "file not found" errors referencing expected cert paths.

### Pitfall 4: ESM vs CJS Module Confusion
**What goes wrong:** Import/require errors when mixing module systems.
**Why it happens:** Node.js determines module type from `package.json` `"type"` field. `.ts` files follow the same rules as `.js` files.
**How to avoid:** Set `"type": "module"` in `package.json` and use `import` syntax consistently. Use `node:` prefix for built-in modules (`node:https`, `node:fs`, `node:path`).
**Warning signs:** "Cannot use import statement outside a module" or "require is not defined" errors.

### Pitfall 5: Path Traversal in Static File Serving
**What goes wrong:** A request like `GET /../../etc/passwd` could read arbitrary files.
**Why it happens:** Naive path joining without validation.
**How to avoid:** Resolve the full path, then verify it starts with the intended static directory. Use `path.resolve()` and check with `startsWith()`.
**Warning signs:** Being able to access files outside the `addin/` directory.

### Pitfall 6: WKWebView Certificate Trust
**What goes wrong:** The Office add-in's WKWebView rejects the WSS connection despite Safari showing no warnings.
**Why it happens:** WKWebView uses the system trust store but may have additional restrictions in sandboxed contexts.
**How to avoid:** Verify that `mkcert -install` completed successfully (CA appears in Keychain Access as trusted). Test the HTTPS URL in Safari first, then in the actual add-in. If issues persist, check that the CA root certificate is in the System keychain (not just login keychain).
**Warning signs:** WebSocket connection fails with a security error in the add-in, but works in Safari.

## Code Examples

### Complete mkcert Setup Script

```bash
# Source: mkcert GitHub README + verified flags
# One-time: install mkcert and create local CA
brew install mkcert
mkcert -install  # Prompts for password on macOS 11.3+

# Per-project: generate certs with predictable filenames
mkdir -p certs
mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1
```

### Minimal HTTPS + WSS Server (TypeScript, no build step)

```typescript
// server/index.ts
// Source: Node.js https docs + ws GitHub README
// Run with: node server/index.ts

import { createServer, type IncomingMessage, type ServerResponse } from 'node:https';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { WebSocketServer, type WebSocket } from 'ws';

const PORT = 8443;
const CERT_PATH = './certs/localhost.pem';
const KEY_PATH = './certs/localhost-key.pem';
const STATIC_DIR = resolve('./addin');

// Check certs exist before starting
if (!existsSync(CERT_PATH) || !existsSync(KEY_PATH)) {
  console.error('TLS certificates not found. Run:');
  console.error('  mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1');
  process.exit(1);
}

const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const urlPath = req.url === '/' ? '/index.html' : req.url ?? '/index.html';
  const filePath = resolve(join(STATIC_DIR, urlPath));

  // Security: prevent path traversal
  if (!filePath.startsWith(STATIC_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  if (!existsSync(filePath)) {
    res.writeHead(404);
    res.end('Not Found');
    return;
  }

  const ext = extname(filePath);
  const contentType = MIME_TYPES[ext] ?? 'application/octet-stream';
  const content = readFileSync(filePath);
  res.writeHead(200, { 'Content-Type': contentType });
  res.end(content);
}

const server = createServer(
  {
    cert: readFileSync(CERT_PATH),
    key: readFileSync(KEY_PATH),
  },
  serveStatic,
);

const wss = new WebSocketServer({ server });

wss.on('connection', (ws: WebSocket) => {
  console.log('WebSocket client connected');
  ws.on('message', (data: Buffer) => {
    console.log('Received:', data.toString());
  });
  ws.on('close', () => {
    console.log('WebSocket client disconnected');
  });
});

server.listen(PORT, () => {
  console.log(`HTTPS server: https://localhost:${PORT}`);
  console.log(`WSS server:   wss://localhost:${PORT}`);
});
```

### tsconfig.json (for editor + type checking only)

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "noEmit": true,
    "strict": true,
    "erasableSyntaxOnly": true,
    "verbatimModuleSyntax": true,
    "skipLibCheck": true,
    "resolveJsonModule": true,
    "isolatedModules": true
  },
  "include": ["server/**/*.ts"],
  "exclude": ["node_modules"]
}
```

**Key options explained:**
- `noEmit: true` -- Node.js runs .ts directly; tsc is only for checking
- `erasableSyntaxOnly: true` -- Errors on enums/decorators that Node.js can't handle
- `verbatimModuleSyntax: true` -- Forces `import type` for type-only imports
- `module: "NodeNext"` -- Matches Node.js ESM resolution

### package.json

```json
{
  "name": "powerpoint-bridge",
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "start": "node server/index.ts",
    "typecheck": "tsc --noEmit",
    "setup-certs": "mkdir -p certs && mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1"
  },
  "dependencies": {
    "ws": "^8.19.0"
  },
  "devDependencies": {
    "@types/ws": "^8.18.1",
    "@types/node": "^22.0.0",
    "typescript": "^5.9.0"
  }
}
```

### WebSocket Test Client (for verifying WSS works)

```bash
# Using Node.js native WebSocket client (available in Node 24)
node -e "
const ws = new WebSocket('wss://localhost:8443');
ws.onopen = () => { console.log('Connected'); ws.send('hello'); };
ws.onmessage = (e) => { console.log('Received:', e.data); ws.close(); };
ws.onerror = (e) => console.error('Error:', e.message);
"
```

Note: The native WebSocket client in Node.js may reject self-signed certs. If so, use:
```bash
NODE_TLS_REJECT_UNAUTHORIZED=0 node -e "..."
```
Or use the `ws` library client which can accept custom CA:
```bash
node -e "
import { WebSocket } from 'ws';
const ws = new WebSocket('wss://localhost:8443', { rejectUnauthorized: false });
ws.on('open', () => { console.log('Connected'); ws.send('hello'); });
ws.on('message', (data) => console.log('Received:', data.toString()));
"
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| ts-node / tsx for running TS | Node.js 24 native type stripping | v23.6.0 (late 2024), stable in v24.3.0 | No dev dependency needed for TS execution |
| OpenSSL for dev certs | mkcert | 2018+ (mature) | Eliminates manual Keychain trust steps |
| Express for any HTTP server | Native `node:https` for simple use cases | Always was an option | Fewer dependencies for simple static serving |
| Separate WS port | ws attaches to existing HTTP server | ws has supported this for years | Single port for HTTPS + WSS |

**Deprecated/outdated:**
- `ts-node`: Has ESM compatibility issues; tsx or native Node.js TS is preferred
- Manual `openssl req` commands: Replaced by mkcert for localhost development
- `WebSocket` polyfills for Node.js: Native WebSocket client available since Node 21 (stable in 22.4.0)

## Open Questions

1. **WKWebView trust of mkcert CA in PowerPoint sandbox**
   - What we know: mkcert CA is installed in macOS system trust store; Safari trusts it; WKWebView generally uses the same trust store
   - What's unclear: Whether PowerPoint's WKWebView sandbox has additional certificate restrictions
   - Recommendation: Build and test in Phase 1. If WSS fails from the add-in, the error will be visible. Fallback: add the mkcert root CA to the PowerPoint container's trust store

2. **Port number choice**
   - What we know: Must be a high port (>1024). Common choices: 3000, 8080, 8443
   - What's unclear: Whether any other local service conflicts on the user's machine
   - Recommendation: Use 8443 (standard alternative HTTPS port). Make it configurable via environment variable.

3. **Native WebSocket client TLS behavior in Node 24**
   - What we know: Node 24 has native `WebSocket` global (client only). It uses the system trust store.
   - What's unclear: Whether the native WS client trusts mkcert's CA for test scripts, or needs `NODE_TLS_REJECT_UNAUTHORIZED=0`
   - Recommendation: Test during implementation. Provide both native and `ws`-library test options.

## Sources

### Primary (HIGH confidence)

- [Node.js TypeScript documentation](https://nodejs.org/api/typescript.html) - Native TS support details, erasable syntax, limitations
- [ws GitHub repository](https://github.com/websockets/ws) - HTTPS+WSS integration pattern, API
- [mkcert GitHub repository](https://github.com/FiloSottile/mkcert) - Installation, flags, CA trust behavior
- Verified on local machine: Node.js 24.8.0 native TS, ws 8.19.0, macOS 26.2

### Secondary (MEDIUM confidence)

- [mkcert issue #415](https://github.com/FiloSottile/mkcert/issues/415) - macOS Big Sur 11.3+ password requirement
- [Office Add-in mkcert guide](https://kags.me.ke/posts/office-add-in-mkcert-localhost-ssl-certificate/) - mkcert for Office add-in development
- [Node.js running TypeScript natively guide](https://nodejs.org/en/learn/typescript/run-natively) - Usage patterns
- [MDN Node.js server without framework](https://developer.mozilla.org/en-US/docs/Learn/Server-side/Node_server_without_framework) - Static file serving patterns

### Tertiary (LOW confidence)

- WebSearch results about WKWebView + mkcert trust (multiple sources agree but no authoritative Office-specific confirmation)

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - All versions verified on local machine; ws and mkcert are well-established
- Architecture: HIGH - Single-process HTTPS+WSS pattern is documented in ws README and widely used
- TypeScript setup: HIGH - Native TS support tested directly on Node.js 24.8.0 on this machine
- Pitfalls: MEDIUM - mkcert macOS issues verified via GitHub issues; WKWebView trust is inferred from system behavior
- WKWebView certificate trust: LOW - No authoritative source confirms behavior inside PowerPoint's sandbox specifically

**Research date:** 2026-02-06
**Valid until:** 2026-03-06 (stable infrastructure, unlikely to change)
