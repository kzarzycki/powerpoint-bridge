import { createServer } from 'node:https';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { WebSocketServer } from 'ws';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { WebSocket } from 'ws';
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
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
});

// ---------------------------------------------------------------------------
// MCP server (stdio transport — coexists with HTTPS+WSS on port 8443)
// ---------------------------------------------------------------------------

const mcpServer = new McpServer({
  name: "powerpoint-bridge",
  version: "0.1.0",
});

// --- Tool 1: get_presentation ---
mcpServer.tool(
  "get_presentation",
  "Returns the structure of the currently open PowerPoint presentation including all slides with their IDs and shape summaries (count, names, types). Use this first to understand what's in the presentation before making changes.",
  async () => {
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
      const target = resolveTarget();
      const result = await sendCommand('executeCode', { code }, target.ws);
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
    }
  }
);

// --- Tool 2: get_slide ---
mcpServer.tool(
  "get_slide",
  "Returns detailed information about all shapes on a specific slide, including text content, positions (left, top in points), sizes (width, height in points), and fill colors. Use slideIndex from get_presentation results (zero-based).",
  { slideIndex: z.number().int().min(0).describe("Zero-based slide index from get_presentation results") },
  async ({ slideIndex }) => {
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
      const target = resolveTarget();
      const result = await sendCommand('executeCode', { code }, target.ws);
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
    }
  }
);

// --- Tool 3: execute_officejs ---
mcpServer.tool(
  "execute_officejs",
  "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. For positioning, all values are in points (1 point = 1/72 inch). Common operations: add shapes, set text, change colors, add/delete slides.",
  { code: z.string().describe("Office.js code to execute. Runs inside PowerPoint.run() with 'context' available. Use 'return' to send back a result.") },
  async ({ code }) => {
    try {
      const target = resolveTarget();
      const result = await sendCommand('executeCode', { code }, target.ws);
      return { content: [{ type: "text" as const, text: JSON.stringify(result ?? { success: true }, null, 2) }] };
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      return { content: [{ type: "text" as const, text: "Error: " + message }], isError: true };
    }
  }
);

const transport = new StdioServerTransport();
await mcpServer.connect(transport);
console.error("MCP server connected via stdio");
