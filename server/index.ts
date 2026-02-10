import { randomUUID } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { createServer as createHttpServer } from 'node:http'
import { createServer } from 'node:https'
import { extname, join, resolve } from 'node:path'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js'
import type { WebSocket } from 'ws'
import { WebSocketServer } from 'ws'

import { ConnectionPool } from './bridge.ts'
import { clearSessionWarnings, registerTools } from './tools.ts'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const PORT = 8443
const MCP_PORT = 3001
const CERT_PATH = './certs/localhost.pem'
const KEY_PATH = './certs/localhost-key.pem'
const STATIC_DIR = resolve('./addin')

// ---------------------------------------------------------------------------
// Startup cert check
// ---------------------------------------------------------------------------

if (!existsSync(CERT_PATH) || !existsSync(KEY_PATH)) {
  console.error(
    'Error: TLS certificate files not found.\n' +
      `  Expected: ${CERT_PATH} and ${KEY_PATH}\n` +
      '  Run: npm run setup-certs',
  )
  process.exit(1)
}

// ---------------------------------------------------------------------------
// Shared state
// ---------------------------------------------------------------------------

const pool = new ConnectionPool()
const mcpTransports = new Map<string, StreamableHTTPServerTransport>()

// ---------------------------------------------------------------------------
// MIME type map
// ---------------------------------------------------------------------------

const MIME_TYPES = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
} as const

type KnownExt = keyof typeof MIME_TYPES

function getMimeType(filePath: string): string {
  const ext = extname(filePath) as KnownExt
  return MIME_TYPES[ext] ?? 'application/octet-stream'
}

// ---------------------------------------------------------------------------
// MCP HTTP transport helpers
// ---------------------------------------------------------------------------

function parseJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    req.on('data', (chunk: Buffer) => chunks.push(chunk))
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString()))
      } catch {
        reject(new Error('Invalid JSON body'))
      }
    })
    req.on('error', reject)
  })
}

function createMcpSession(transport: StreamableHTTPServerTransport): McpServer {
  const server = new McpServer({
    name: 'powerpoint-bridge',
    version: '0.1.0',
  })
  registerTools(
    server,
    pool,
    () => transport.sessionId ?? undefined,
    () => mcpTransports.size,
  )
  return server
}

async function handleMcpPost(req: IncomingMessage, res: ServerResponse): Promise<void> {
  try {
    const body = await parseJsonBody(req)
    const sessionId = req.headers['mcp-session-id'] as string | undefined

    if (sessionId && mcpTransports.has(sessionId)) {
      await mcpTransports.get(sessionId)!.handleRequest(req, res, body)
    } else if (!sessionId && isInitializeRequest(body)) {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          mcpTransports.set(sid, transport)
          console.error(`MCP session initialized: ${sid}`)
        },
      })
      transport.onclose = () => {
        const sid = transport.sessionId
        if (sid) {
          mcpTransports.delete(sid)
          clearSessionWarnings(sid)
          console.error(`MCP session closed: ${sid}`)
        }
      }
      const server = createMcpSession(transport)
      await server.connect(transport)
      await transport.handleRequest(req, res, body)
    } else {
      res.writeHead(400, { 'Content-Type': 'application/json' })
      res.end(
        JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Bad request: no valid session' }, id: null }),
      )
    }
  } catch (err) {
    console.error('MCP POST error:', err)
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32603, message: 'Internal error' }, id: null }))
    }
  }
}

async function handleMcpGet(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (!sessionId || !mcpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' })
    res.end(
      JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }),
    )
    return
  }
  await mcpTransports.get(sessionId)!.handleRequest(req, res)
}

async function handleMcpDelete(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (!sessionId || !mcpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' })
    res.end(
      JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }),
    )
    return
  }
  await mcpTransports.get(sessionId)!.handleRequest(req, res)
}

// ---------------------------------------------------------------------------
// Static file handler
// ---------------------------------------------------------------------------

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const rawUrl = (req.url ?? '/').split('?')[0]

  if (rawUrl === '/api/test') {
    let target: ReturnType<ConnectionPool['resolveTarget']>
    try {
      target = pool.resolveTarget()
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err)
      res.writeHead(500, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify({ error: message }))
      return
    }
    pool
      .sendCommand(
        'executeCode',
        {
          code: 'var c = context.presentation.slides.getCount(); await context.sync(); return c.value;',
        },
        target.ws,
      )
      .then((result) => {
        res.writeHead(200, { 'Content-Type': 'application/json' })
        res.end(JSON.stringify({ slideCount: result }))
      })
      .catch((err: Error) => {
        res.writeHead(500, { 'Content-Type': 'application/json' })
        res.end(JSON.stringify({ error: err.message }))
      })
    return
  }

  const urlPath = rawUrl === '/' ? '/index.html' : rawUrl
  const filePath = resolve(join(STATIC_DIR, urlPath))

  if (!filePath.startsWith(STATIC_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' })
    res.end('403 Forbidden')
    return
  }

  if (!existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('404 Not Found')
    return
  }

  const content = readFileSync(filePath)
  res.writeHead(200, { 'Content-Type': getMimeType(filePath) })
  res.end(content)
}

// ---------------------------------------------------------------------------
// HTTPS server
// ---------------------------------------------------------------------------

const cert = readFileSync(CERT_PATH)
const key = readFileSync(KEY_PATH)

const server = createServer({ cert, key }, serveStatic)

// ---------------------------------------------------------------------------
// WebSocket server
// ---------------------------------------------------------------------------

const wss = new WebSocketServer({ server })

wss.on('connection', (ws: WebSocket) => {
  console.error('WebSocket client connected')

  ws.on('message', (data: Buffer) => {
    let msg: { type?: string; id?: string; data?: unknown; error?: { message?: string }; documentUrl?: string }
    try {
      msg = JSON.parse(data.toString())
    } catch {
      console.error('Invalid JSON from add-in:', data.toString())
      return
    }

    if ((msg.type === 'response' || msg.type === 'error') && msg.id) {
      pool.handleResponse(msg.id, msg.type, msg.data, msg.error?.message)
    }

    if (msg.type === 'ready') {
      const documentUrl = typeof msg.documentUrl === 'string' && msg.documentUrl.length > 0 ? msg.documentUrl : null
      const presentationId = pool.generateId(documentUrl)
      pool.add(presentationId, {
        ws,
        ready: true,
        presentationId,
        filePath: documentUrl,
      })
      console.error(`Add-in ready: ${presentationId}`)
    }
  })

  ws.on('close', () => {
    const disconnectedId = pool.removeBySocket(ws)
    if (disconnectedId) {
      console.error(`Add-in disconnected: ${disconnectedId}`)
    }
    pool.rejectPendingForSocket(ws)
  })

  ws.on('error', (err: Error) => {
    console.error('WebSocket error:', err.message)
  })
})

// ---------------------------------------------------------------------------
// Start HTTPS server
// ---------------------------------------------------------------------------

server.listen(PORT, () => {
  console.error('Bridge server running')
  console.error(`  HTTPS: https://localhost:${PORT}`)
  console.error(`  WSS:   wss://localhost:${PORT}`)
})

// ---------------------------------------------------------------------------
// Plain HTTP server for MCP
// ---------------------------------------------------------------------------

function handleMcpRequest(req: IncomingMessage, res: ServerResponse): void {
  const url = (req.url ?? '/').split('?')[0]

  if (url === '/mcp') {
    if (req.method === 'POST') {
      handleMcpPost(req, res)
    } else if (req.method === 'GET') {
      handleMcpGet(req, res)
    } else if (req.method === 'DELETE') {
      handleMcpDelete(req, res)
    } else {
      res.writeHead(405)
      res.end()
    }
    return
  }

  res.writeHead(404, { 'Content-Type': 'text/plain' })
  res.end('MCP endpoint is at /mcp')
}

const mcpHttpServer = createHttpServer(handleMcpRequest)
mcpHttpServer.listen(MCP_PORT, () => {
  console.error(`  MCP:   http://localhost:${MCP_PORT}/mcp`)
})
