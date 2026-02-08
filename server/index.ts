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
}

const pendingRequests = new Map<string, PendingRequest>();
const COMMAND_TIMEOUT = 30_000;
let addinClient: WebSocket | null = null;
let addinReady = false;

function sendCommand(action: string, params: Record<string, unknown>): Promise<unknown> {
  if (!addinClient || addinClient.readyState !== 1) {
    return Promise.reject(new Error('Add-in not connected'));
  }
  if (!addinReady) {
    return Promise.reject(new Error('Add-in not ready'));
  }

  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      reject(new Error(`Command timed out after ${COMMAND_TIMEOUT}ms`));
    }, COMMAND_TIMEOUT);

    pendingRequests.set(id, { resolve, reject, timer });
    addinClient!.send(JSON.stringify({ type: 'command', id, action, params }));
  });
}

// ---------------------------------------------------------------------------
// Static file handler
// ---------------------------------------------------------------------------

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  // Strip query string — Office.js/WKWebView appends ?_host_Info=... params
  const rawUrl = (req.url ?? '/').split('?')[0];

  // API: test endpoint — sends a slide count command to the add-in
  if (rawUrl === '/api/test') {
    sendCommand('executeCode', {
      code: 'var c = context.presentation.slides.getCount(); await context.sync(); return c.value;',
    })
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
  addinClient = ws;
  console.error('WebSocket client connected');

  ws.on('message', (data: Buffer) => {
    let msg: { type?: string; id?: string; data?: unknown; error?: { message?: string } };
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
      addinReady = true;
      console.error('Add-in ready to receive commands');
    }
  });

  ws.on('close', () => {
    for (const [id, pending] of pendingRequests) {
      clearTimeout(pending.timer);
      pending.reject(new Error('Add-in disconnected'));
    }
    pendingRequests.clear();
    addinClient = null;
    addinReady = false;
    console.error('WebSocket client disconnected');
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
      const result = await sendCommand('executeCode', { code });
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
      const result = await sendCommand('executeCode', { code });
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
      const result = await sendCommand('executeCode', { code });
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
