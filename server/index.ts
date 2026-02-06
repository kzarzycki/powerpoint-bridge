import { createServer } from 'node:https';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { WebSocketServer } from 'ws';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { WebSocket } from 'ws';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const PORT = 8443;
const CERT_PATH = './certs/localhost.pem';
const KEY_PATH = './certs/localhost-key.pem';
const STATIC_DIR = resolve('./addin');

// ---------------------------------------------------------------------------
// Startup cert check
// ---------------------------------------------------------------------------

if (!existsSync(CERT_PATH) || !existsSync(KEY_PATH)) {
  console.error(
    'Error: TLS certificate files not found.\n' +
    `  Expected: ${CERT_PATH} and ${KEY_PATH}\n` +
    '  Run: npm run setup-certs'
  );
  process.exit(1);
}

// ---------------------------------------------------------------------------
// MIME type map (as const — no enums in Node 24 type stripping)
// ---------------------------------------------------------------------------

const MIME_TYPES = {
  '.html': 'text/html; charset=UTF-8',
  '.js':   'text/javascript',
  '.css':  'text/css',
  '.json': 'application/json',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
} as const;

type KnownExt = keyof typeof MIME_TYPES;

function getMimeType(filePath: string): string {
  const ext = extname(filePath) as KnownExt;
  return MIME_TYPES[ext] ?? 'application/octet-stream';
}

// ---------------------------------------------------------------------------
// Static file handler
// ---------------------------------------------------------------------------

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const urlPath = req.url === '/' ? '/index.html' : (req.url ?? '/index.html');

  const filePath = resolve(join(STATIC_DIR, urlPath));

  // Security: prevent path traversal
  if (!filePath.startsWith(STATIC_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('403 Forbidden');
    return;
  }

  if (!existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('404 Not Found');
    return;
  }

  const content = readFileSync(filePath);
  res.writeHead(200, { 'Content-Type': getMimeType(filePath) });
  res.end(content);
}

// ---------------------------------------------------------------------------
// HTTPS server
// ---------------------------------------------------------------------------

const cert = readFileSync(CERT_PATH);
const key = readFileSync(KEY_PATH);

const server = createServer({ cert, key }, serveStatic);

// ---------------------------------------------------------------------------
// WebSocket server (shares port with HTTPS via { server } option)
// ---------------------------------------------------------------------------

const wss = new WebSocketServer({ server });

wss.on('connection', (ws: WebSocket) => {
  console.log('WebSocket client connected');

  ws.on('message', (data: Buffer) => {
    console.log('WebSocket message:', data.toString());
  });

  ws.on('close', () => {
    console.log('WebSocket client disconnected');
  });

  ws.on('error', (err: Error) => {
    console.error('WebSocket error:', err.message);
  });
});

// ---------------------------------------------------------------------------
// Start listening
// ---------------------------------------------------------------------------

server.listen(PORT, () => {
  console.log('Bridge server running');
  console.log(`  HTTPS: https://localhost:${PORT}`);
  console.log(`  WSS:   wss://localhost:${PORT}`);
});
