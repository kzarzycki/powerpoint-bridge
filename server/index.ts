import { createServer } from 'node:https';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { WebSocketServer } from 'ws';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { WebSocket } from 'ws';
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";

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
// Command infrastructure
// ---------------------------------------------------------------------------

interface PendingRequest {
  resolve: (data: unknown) => void;
  reject: (err: Error) => void;
  timer: ReturnType<typeof setTimeout>;
  ws: WebSocket;
}

const pendingRequests = new Map<string, PendingRequest>();
const COMMAND_TIMEOUT = 30_000;

// ---------------------------------------------------------------------------
// Multi-connection pool
// ---------------------------------------------------------------------------

interface AddinConnection {
  ws: WebSocket;
  ready: boolean;
  presentationId: string;
  filePath: string | null;
}

const addinConnections = new Map<string, AddinConnection>();
let untitledCounter = 0;

function sendCommand(action: string, params: Record<string, unknown>, targetWs: WebSocket): Promise<unknown> {
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      reject(new Error(`Command timed out after ${COMMAND_TIMEOUT}ms`));
    }, COMMAND_TIMEOUT);

    pendingRequests.set(id, { resolve, reject, timer, ws: targetWs });
    targetWs.send(JSON.stringify({ type: 'command', id, action, params }));
  });
}

function resolveTarget(presentationId?: string): AddinConnection {
  if (addinConnections.size === 0) {
    throw new Error('No presentations connected. Open a PowerPoint file with the bridge add-in loaded.');
  }
  if (presentationId) {
    const conn = addinConnections.get(presentationId);
    if (!conn) throw new Error('Presentation not found: ' + presentationId + '. Use list_presentations to see connected presentations.');
    if (!conn.ready) throw new Error('Presentation connected but not ready: ' + presentationId);
    return conn;
  }
  if (addinConnections.size === 1) {
    const single = addinConnections.values().next().value!;
    if (!single.ready) throw new Error('Add-in connected but not ready');
    return single;
  }
  const ids = [...addinConnections.keys()];
  throw new Error('Multiple presentations connected. Specify presentationId parameter. Available: ' + ids.join(', '));
}

// ---------------------------------------------------------------------------
// MCP session tracking
// ---------------------------------------------------------------------------

const mcpTransports = new Map<string, StreamableHTTPServerTransport>();

// Track which sessions have been warned about concurrent access per presentation
const sessionConcurrentWarnings = new Map<string, Set<string>>();

// ---------------------------------------------------------------------------
// Concurrent access warning helper
// ---------------------------------------------------------------------------

function getConcurrentWarning(mcpSessionId: string | undefined, presentationId: string): string | null {
  if (!mcpSessionId) return null;
  if (mcpTransports.size <= 1) return null;

  const warned = sessionConcurrentWarnings.get(mcpSessionId);
  if (warned?.has(presentationId)) return null;

  if (!warned) {
    sessionConcurrentWarnings.set(mcpSessionId, new Set([presentationId]));
  } else {
    warned.add(presentationId);
  }

  return '\n\nNote: Other MCP sessions are also connected to the bridge. If they target this presentation, changes apply immediately (last-write-wins).';
}

// ---------------------------------------------------------------------------
// Tool registration (per-session factory)
// ---------------------------------------------------------------------------

function registerTools(server: McpServer, getSessionId: () => string | undefined): void {

  // --- Tool: list_presentations ---
  server.tool(
    "list_presentations",
    "Lists all PowerPoint presentations currently connected to the bridge server. Shows presentation IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the presentationId to pass to other tools when multiple presentations are open.",
    async () => {
      const presentations = [];
      for (const [id, conn] of addinConnections) {
        presentations.push({
          presentationId: id,
          filePath: conn.filePath,
          ready: conn.ready,
        });
      }
      return {
        content: [{
          type: "text" as const,
          text: presentations.length === 0
            ? "No presentations connected. Open a PowerPoint file with the bridge add-in loaded."
            : JSON.stringify(presentations, null, 2)
        }]
      };
    }
  );

  // --- Tool: get_presentation ---
  server.tool(
    "get_presentation",
    "Returns the structure of the currently open PowerPoint presentation including all slides with their IDs and shape summaries (count, names, types). Use this first to understand what's in the presentation before making changes.",
    { presentationId: z.string().optional().describe("Target presentation ID from list_presentations. Optional when only one presentation is connected.") },
    async ({ presentationId }) => {
      try {
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          for (var i = 0; i < slides.items.length; i++) {
            slides.items[i].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              shapes.push({ name: s.name, type: s.type, id: s.id });
            }
            output.push({ index: i, id: slide.id, shapeCount: shapes.length, shapes: shapes });
          }
          return output;
        `;
        const target = resolveTarget(presentationId);
        const result = await sendCommand('executeCode', { code }, target.ws);
        const warning = getConcurrentWarning(getSessionId(), target.presentationId);
        const text = JSON.stringify(result, null, 2) + (warning ?? '');
        return { content: [{ type: "text" as const, text }] };
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
      }
    }
  );

  // --- Tool: get_slide ---
  server.tool(
    "get_slide",
    "Returns detailed information about all shapes on a specific slide, including text content, positions (left, top in points), sizes (width, height in points), and fill colors. Use slideIndex from get_presentation results (zero-based).",
    {
      slideIndex: z.number().int().min(0).describe("Zero-based slide index from get_presentation results"),
      presentationId: z.string().optional().describe("Target presentation ID from list_presentations. Optional when only one presentation is connected.")
    },
    async ({ slideIndex, presentationId }) => {
      try {
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          slide.shapes.load("items");
          await context.sync();
          var shapes = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var info = {
              name: s.name,
              type: s.type,
              id: s.id,
              left: s.left,
              top: s.top,
              width: s.width,
              height: s.height
            };
            try {
              s.textFrame.load("textRange");
              await context.sync();
              info.text = s.textFrame.textRange.text;
            } catch (e) {
              // Shape has no text frame (e.g., images, connectors)
            }
            try {
              s.fill.load("foregroundColor,type");
              await context.sync();
              info.fill = { type: s.fill.type, color: s.fill.foregroundColor };
            } catch (e) {
              // Shape has no fill or fill not accessible
            }
            shapes.push(info);
          }
          return { slideIndex: ${slideIndex}, slideId: slide.id, shapes: shapes };
        `;
        const target = resolveTarget(presentationId);
        const result = await sendCommand('executeCode', { code }, target.ws);
        const warning = getConcurrentWarning(getSessionId(), target.presentationId);
        const text = JSON.stringify(result, null, 2) + (warning ?? '');
        return { content: [{ type: "text" as const, text }] };
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
      }
    }
  );

  // --- Tool: execute_officejs ---
  server.tool(
    "execute_officejs",
    "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. For positioning, all values are in points (1 point = 1/72 inch). Common operations: add shapes, set text, change colors, add/delete slides.",
    {
      code: z.string().describe("Office.js code to execute. Runs inside PowerPoint.run() with 'context' available. Use 'return' to send back a result."),
      presentationId: z.string().optional().describe("Target presentation ID from list_presentations. Optional when only one presentation is connected.")
    },
    async ({ code, presentationId }) => {
      try {
        const target = resolveTarget(presentationId);
        const result = await sendCommand('executeCode', { code }, target.ws);
        const warning = getConcurrentWarning(getSessionId(), target.presentationId);
        const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '');
        return { content: [{ type: "text" as const, text }] };
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
      }
    }
  );
}

// ---------------------------------------------------------------------------
// MCP HTTP transport helpers
// ---------------------------------------------------------------------------

function parseJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on('data', (chunk: Buffer) => chunks.push(chunk));
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString()));
      } catch {
        reject(new Error('Invalid JSON body'));
      }
    });
    req.on('error', reject);
  });
}

function createMcpSession(transport: StreamableHTTPServerTransport): McpServer {
  const server = new McpServer({
    name: "powerpoint-bridge",
    version: "0.2.0",
  });
  registerTools(server, () => transport.sessionId ?? undefined);
  return server;
}

async function handleMcpPost(req: IncomingMessage, res: ServerResponse): Promise<void> {
  try {
    const body = await parseJsonBody(req);
    const sessionId = req.headers['mcp-session-id'] as string | undefined;

    if (sessionId && mcpTransports.has(sessionId)) {
      await mcpTransports.get(sessionId)!.handleRequest(req, res, body);
    } else if (!sessionId && isInitializeRequest(body)) {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          mcpTransports.set(sid, transport);
          console.error('MCP session initialized: ' + sid);
        },
      });
      transport.onclose = () => {
        const sid = transport.sessionId;
        if (sid) {
          mcpTransports.delete(sid);
          sessionConcurrentWarnings.delete(sid);
          console.error('MCP session closed: ' + sid);
        }
      };
      const server = createMcpSession(transport);
      await server.connect(transport);
      await transport.handleRequest(req, res, body);
    } else {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Bad request: no valid session' }, id: null }));
    }
  } catch (err) {
    console.error('MCP POST error:', err);
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32603, message: 'Internal error' }, id: null }));
    }
  }
}

async function handleMcpGet(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;
  if (!sessionId || !mcpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }));
    return;
  }
  await mcpTransports.get(sessionId)!.handleRequest(req, res);
}

async function handleMcpDelete(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;
  if (!sessionId || !mcpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }));
    return;
  }
  await mcpTransports.get(sessionId)!.handleRequest(req, res);
}

// ---------------------------------------------------------------------------
// Static file handler
// ---------------------------------------------------------------------------

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  // Strip query string — Office.js/WKWebView appends ?_host_Info=... params
  const rawUrl = (req.url ?? '/').split('?')[0];

  // API: test endpoint — sends a slide count command to the add-in
  if (rawUrl === '/api/test') {
    let target: AddinConnection;
    try {
      target = resolveTarget();
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: message }));
      return;
    }
    sendCommand('executeCode', {
      code: 'var c = context.presentation.slides.getCount(); await context.sync(); return c.value;',
    }, target.ws)
      .then((result) => {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ slideCount: result }));
      })
      .catch((err: Error) => {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      });
    return;
  }

  const urlPath = rawUrl === '/' ? '/index.html' : rawUrl;

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
// HTTPS server (with /mcp route for MCP HTTP transport)
// ---------------------------------------------------------------------------

const cert = readFileSync(CERT_PATH);
const key = readFileSync(KEY_PATH);

function handleHttpRequest(req: IncomingMessage, res: ServerResponse): void {
  const url = (req.url ?? '/').split('?')[0];

  if (url === '/mcp') {
    if (req.method === 'POST') { handleMcpPost(req, res); }
    else if (req.method === 'GET') { handleMcpGet(req, res); }
    else if (req.method === 'DELETE') { handleMcpDelete(req, res); }
    else { res.writeHead(405); res.end(); }
    return;
  }

  serveStatic(req, res);
}

const server = createServer({ cert, key }, handleHttpRequest);

// ---------------------------------------------------------------------------
// WebSocket server (shares port with HTTPS via { server } option)
// ---------------------------------------------------------------------------

const wss = new WebSocketServer({ server });

wss.on('connection', (ws: WebSocket) => {
  console.error('WebSocket client connected');

  ws.on('message', (data: Buffer) => {
    let msg: { type?: string; id?: string; data?: unknown; error?: { message?: string }; documentUrl?: string };
    try {
      msg = JSON.parse(data.toString());
    } catch {
      console.error('Invalid JSON from add-in:', data.toString());
      return;
    }

    if ((msg.type === 'response' || msg.type === 'error') && msg.id) {
      const pending = pendingRequests.get(msg.id);
      if (pending) {
        clearTimeout(pending.timer);
        pendingRequests.delete(msg.id);
        if (msg.type === 'response') {
          pending.resolve(msg.data);
        } else {
          pending.reject(new Error(msg.error?.message || 'Command failed'));
        }
      }
    }

    if (msg.type === 'ready') {
      const documentUrl = typeof msg.documentUrl === 'string' && msg.documentUrl.length > 0
        ? msg.documentUrl
        : null;
      const presentationId = documentUrl ?? ('untitled-' + (++untitledCounter));
      const conn: AddinConnection = {
        ws,
        ready: true,
        presentationId,
        filePath: documentUrl,
      };
      addinConnections.set(presentationId, conn);
      console.error('Add-in ready: ' + presentationId);
    }
  });

  ws.on('close', () => {
    // Find which connection disconnected
    let disconnectedId: string | null = null;
    for (const [id, conn] of addinConnections) {
      if (conn.ws === ws) {
        disconnectedId = id;
        break;
      }
    }
    if (disconnectedId) {
      addinConnections.delete(disconnectedId);
      console.error('Add-in disconnected: ' + disconnectedId);
    }
    // Reject only pending requests that were sent to this WebSocket
    for (const [id, pending] of pendingRequests) {
      if (pending.ws === ws) {
        clearTimeout(pending.timer);
        pending.reject(new Error('Add-in disconnected'));
        pendingRequests.delete(id);
      }
    }
  });

  ws.on('error', (err: Error) => {
    console.error('WebSocket error:', err.message);
  });
});

// ---------------------------------------------------------------------------
// Start listening
// ---------------------------------------------------------------------------

server.listen(PORT, () => {
  console.error('Bridge server running');
  console.error(`  HTTPS: https://localhost:${PORT}`);
  console.error(`  WSS:   wss://localhost:${PORT}`);
  console.error(`  MCP:   https://localhost:${PORT}/mcp`);
});
