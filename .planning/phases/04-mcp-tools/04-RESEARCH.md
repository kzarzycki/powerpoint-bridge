# Phase 4: MCP Tools - Research

**Researched:** 2026-02-08
**Domain:** Model Context Protocol (MCP) SDK for TypeScript/Node.js
**Confidence:** HIGH

## Summary

Phase 4 adds an MCP server to the existing Node.js process so that Claude Code can discover and invoke tools (`get_presentation`, `get_slide`, `execute_officejs`) via JSON-RPC over stdio. The official TypeScript SDK (`@modelcontextprotocol/sdk` v1.26.0) provides `McpServer` and `StdioServerTransport` classes that handle all protocol details. Tools are registered with `registerTool()` using Zod schemas for input validation, and return results in `{ content: [{ type: "text", text: "..." }] }` format.

The critical integration challenge is that MCP's stdio transport uses `process.stdout` for JSON-RPC messages, meaning every `console.log()` in the existing server code must be replaced with `console.error()` (which writes to stderr). The HTTPS+WSS server on port 8443 and MCP stdio are completely independent I/O channels and coexist without conflict in a single Node.js process.

**Primary recommendation:** Install `@modelcontextprotocol/sdk` and `zod`, replace all `console.log` with `console.error`, create an `McpServer` with `StdioServerTransport`, and register three tools that compose Office.js code strings and dispatch them via the existing `sendCommand('executeCode', ...)` function.

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| `@modelcontextprotocol/sdk` | 1.26.0 | MCP server SDK with stdio transport | Official Anthropic SDK, the only maintained TS implementation |
| `zod` | ^3.25 | Input schema validation for MCP tools | Required peer dependency of the SDK |

### Supporting

No additional libraries needed. The existing `ws` library handles WebSocket, and the SDK handles all MCP protocol concerns.

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `registerTool()` API | `tool()` API (deprecated) | `tool()` still works in v1.26.0 but is marked `@deprecated`. Use `registerTool()` for forward compatibility. |
| `@modelcontextprotocol/sdk` v1.x | v2 (pre-alpha on main branch) | v2 is not released yet (Q1 2026 target). v1.x is production-stable and receives bug fixes. |

**Installation:**
```bash
npm install @modelcontextprotocol/sdk zod
```

## Architecture Patterns

### How MCP Fits Into the Existing Server

```
Claude Code  --stdin-->  [StdioServerTransport]  --calls-->  MCP Tool Handler
                                                                    |
                                                          sendCommand('executeCode', { code })
                                                                    |
                                                          [Existing WSS]  -->  Add-in
                                                                    |
                                                          [Existing WSS]  <--  Response
                                                                    |
Claude Code  <--stdout-- [StdioServerTransport]  <--returns--  MCP Tool Handler


HTTPS+WSS on port 8443 (unchanged, for add-in)
MCP stdio on process.stdin/stdout (new, for Claude Code)
Both coexist in the same Node.js process.
```

### Pattern 1: McpServer with StdioServerTransport

**What:** Create an MCP server and connect it to stdio transport. The SDK reads JSON-RPC from stdin and writes responses to stdout.

**When to use:** Always -- this is the standard pattern for local MCP servers spawned by Claude Code.

**Example:**
```typescript
// Source: Official MCP quickstart (modelcontextprotocol.io/docs/develop/build-server)
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const mcpServer = new McpServer({
  name: "powerpoint-bridge",
  version: "0.1.0",
});

// Register tools here (see Pattern 2)

const transport = new StdioServerTransport();
await mcpServer.connect(transport);
console.error("MCP server connected via stdio");  // stderr, NOT stdout
```

### Pattern 2: Tool Registration with registerTool()

**What:** Register MCP tools using the `registerTool()` method with a config object containing description and optional inputSchema.

**When to use:** For every tool exposed to Claude Code.

**Example -- Tool with no parameters:**
```typescript
// Source: SDK type definitions (mcp.d.ts v1.26.0)
mcpServer.registerTool(
  "get_presentation",
  {
    description: "Get the structure of the current PowerPoint presentation...",
  },
  async () => {
    const result = await sendCommand('executeCode', { code: '...' });
    return {
      content: [{ type: "text", text: JSON.stringify(result) }],
    };
  }
);
```

**Example -- Tool with parameters:**
```typescript
// Source: Official MCP quickstart (modelcontextprotocol.io/docs/develop/build-server)
mcpServer.registerTool(
  "get_slide",
  {
    description: "Get detailed information about a specific slide...",
    inputSchema: {
      slideIndex: z.number().int().min(0).describe("Zero-based slide index"),
    },
  },
  async ({ slideIndex }) => {
    const result = await sendCommand('executeCode', { code: `...${slideIndex}...` });
    return {
      content: [{ type: "text", text: JSON.stringify(result) }],
    };
  }
);
```

### Pattern 3: Tool Error Handling

**What:** Return errors from tools using the `isError` flag in the result object.

**When to use:** When sendCommand fails (add-in disconnected, timeout, Office.js error).

**Example:**
```typescript
mcpServer.registerTool(
  "execute_officejs",
  {
    description: "Execute Office.js code in the live PowerPoint presentation",
    inputSchema: {
      code: z.string().describe("Office.js code to execute inside PowerPoint.run()"),
    },
  },
  async ({ code }) => {
    try {
      const result = await sendCommand('executeCode', { code });
      return {
        content: [{ type: "text", text: JSON.stringify(result ?? { success: true }) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text", text: `Error: ${(err as Error).message}` }],
        isError: true,
      };
    }
  }
);
```

### Pattern 4: Import Paths for ESM

**What:** The SDK uses `.js` extensions in import paths, which is required for Node.js ESM resolution.

**When to use:** Always, in every import from the SDK.

**Correct imports:**
```typescript
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
```

**Incorrect (will fail):**
```typescript
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp";  // No .js = resolution error
```

### Anti-Patterns to Avoid

- **Using console.log() anywhere in server code:** stdout is reserved for MCP JSON-RPC messages. Any stray output corrupts the protocol and crashes the session.
- **Using the deprecated `tool()` method:** It still works in v1.26.0 but is deprecated. Use `registerTool()` for forward compatibility with v2.
- **Creating separate MCP request/response handlers:** The SDK handles all JSON-RPC framing, message parsing, and response routing. Never parse stdin manually.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| JSON-RPC protocol | Custom stdin/stdout JSON parser | `StdioServerTransport` from SDK | Handles framing, buffering, newline delimiters, error responses |
| Tool schema validation | Manual parameter checking | Zod schemas via `registerTool({ inputSchema })` | SDK validates automatically, generates JSON Schema for tool listing |
| Tool discovery | Custom `list_tools` handler | SDK's built-in `ListToolsRequestSchema` handler | Automatically responds with registered tools and their schemas |
| Error serialization | Custom error JSON formatting | Return `{ content: [...], isError: true }` | SDK wraps it in proper JSON-RPC error response |

**Key insight:** The MCP SDK handles ALL protocol concerns. The only custom code needed is the tool handler logic (composing Office.js code and calling `sendCommand`).

## Common Pitfalls

### Pitfall 1: console.log Corrupts MCP Stdio

**What goes wrong:** Any `console.log()` call writes to `process.stdout`, which injects non-JSON-RPC data into the MCP protocol stream. The client (Claude Code) fails to parse the corrupted stream and the server crashes or hangs.
**Why it happens:** The existing server/index.ts has 6 `console.log()` calls for status messages ("Bridge server running", "WebSocket client connected", etc.).
**How to avoid:** Replace ALL `console.log()` with `console.error()` throughout the entire server codebase. `console.error()` writes to stderr, which MCP ignores.
**Warning signs:** "MCP server failed to start", garbled JSON errors in Claude Code logs, server process exits immediately.

**Lines to change in server/index.ts:**
```
Line 149: console.log('WebSocket client connected');
Line 175: console.log('Add-in ready to receive commands');
Line 187: console.log('WebSocket client disconnected');
Line 200: console.log('Bridge server running');
Line 201: console.log(`  HTTPS: https://localhost:${PORT}`);
Line 202: console.log(`  WSS:   wss://localhost:${PORT}`);
```

### Pitfall 2: Forgetting .js Extension in Imports

**What goes wrong:** Node.js ESM resolution requires file extensions. Importing `@modelcontextprotocol/sdk/server/mcp` without `.js` results in ERR_MODULE_NOT_FOUND.
**Why it happens:** TypeScript traditionally doesn't require file extensions, but Node.js ESM does.
**How to avoid:** Always include `.js` in SDK imports.
**Warning signs:** `ERR_MODULE_NOT_FOUND` error on startup.

### Pitfall 3: Tool Handler Not Returning Content Array

**What goes wrong:** MCP tools must return `{ content: [{ type: "text", text: "..." }] }`. Returning a plain string or object causes the SDK to throw an error.
**Why it happens:** The MCP protocol specifies a strict response format with content types (text, image, resource).
**How to avoid:** Always wrap results in the content array format. Use `JSON.stringify()` for objects.
**Warning signs:** Tool calls return empty results or throw "invalid result" errors.

### Pitfall 4: Add-in Not Connected When Tool Is Called

**What goes wrong:** Claude Code calls an MCP tool before the PowerPoint add-in has connected via WebSocket. The existing `sendCommand()` rejects with "Add-in not connected".
**Why it happens:** Claude Code spawns the Node.js process and immediately sends MCP requests. The add-in needs to be open in PowerPoint and connected via WSS.
**How to avoid:** Catch the `sendCommand` error and return a clear, actionable error message via `isError: true`, telling Claude to wait for the add-in or instructing the user to open PowerPoint.
**Warning signs:** Every tool call fails with "Add-in not connected".

### Pitfall 5: Zod Version Mismatch

**What goes wrong:** Installing `zod` outside the `^3.25` range causes the SDK's internal schema conversion to fail.
**Why it happens:** The SDK requires `zod ^3.25 || ^4.0`. Earlier zod v3 versions (e.g., 3.22) are incompatible.
**How to avoid:** Install `zod@3` which resolves to latest 3.x (currently 3.25.76, well within range).
**Warning signs:** Type errors or runtime errors in `zod-to-json-schema` conversion.

## Code Examples

### Complete MCP Server Integration (Skeleton)

```typescript
// Source: Official MCP quickstart + SDK type definitions (v1.26.0)
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// ... existing server code (HTTPS, WSS, sendCommand) ...
// ... all console.log replaced with console.error ...

const mcpServer = new McpServer({
  name: "powerpoint-bridge",
  version: "0.1.0",
});

// Tool 1: get_presentation (no params)
mcpServer.registerTool(
  "get_presentation",
  {
    description: "Returns the structure of the currently open PowerPoint presentation: all slides with their indices, IDs, and shape summaries (count, names, types).",
  },
  async () => {
    try {
      const result = await sendCommand('executeCode', {
        code: `
          var slides = context.presentation.slides;
          slides.load('items');
          await context.sync();
          var output = [];
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            slide.shapes.load('items');
          }
          await context.sync();
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            var shapes = slide.shapes.items.map(function(s) {
              return { name: s.name, type: s.type, id: s.id };
            });
            output.push({ index: i, id: slide.id, shapeCount: shapes.length, shapes: shapes });
          }
          return output;
        `,
      });
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    } catch (err) {
      return { content: [{ type: "text", text: (err as Error).message }], isError: true };
    }
  }
);

// Tool 2: get_slide (takes slideIndex)
mcpServer.registerTool(
  "get_slide",
  {
    description: "Returns detailed information for a single slide: all shapes with text content, positions (left, top), sizes (width, height), fill colors, and font colors.",
    inputSchema: {
      slideIndex: z.number().int().min(0).describe("Zero-based index of the slide to inspect"),
    },
  },
  async ({ slideIndex }) => {
    try {
      const result = await sendCommand('executeCode', {
        code: `
          var slides = context.presentation.slides;
          slides.load('items');
          await context.sync();
          var slide = slides.items[${slideIndex}];
          if (!slide) return { error: 'Slide index ${slideIndex} out of range' };
          slide.shapes.load('items');
          await context.sync();
          var shapes = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            s.load('name,type,id,left,top,width,height');
            if (s.textFrame) {
              s.textFrame.load('textRange');
            }
            if (s.fill) {
              s.fill.load('foregroundColor,type');
            }
          }
          await context.sync();
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var info = {
              name: s.name, type: s.type, id: s.id,
              left: s.left, top: s.top, width: s.width, height: s.height,
            };
            try {
              if (s.textFrame && s.textFrame.textRange) {
                s.textFrame.textRange.load('text');
                await context.sync();
                info.text = s.textFrame.textRange.text;
              }
            } catch(e) {}
            try {
              if (s.fill) {
                info.fillColor = s.fill.foregroundColor;
                info.fillType = s.fill.type;
              }
            } catch(e) {}
            shapes.push(info);
          }
          return { slideIndex: ${slideIndex}, slideId: slide.id, shapes: shapes };
        `,
      });
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    } catch (err) {
      return { content: [{ type: "text", text: (err as Error).message }], isError: true };
    }
  }
);

// Tool 3: execute_officejs (takes code string)
mcpServer.registerTool(
  "execute_officejs",
  {
    description: "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(context => { ... }) and has access to 'context' (the RequestContext) and 'PowerPoint' (the namespace). Use 'await context.sync()' to flush operations. Return a value to get it back.",
    inputSchema: {
      code: z.string().describe("Office.js code to execute. Has access to 'context' and 'PowerPoint'. Use 'await context.sync()' and 'return' for results."),
    },
  },
  async ({ code }) => {
    try {
      const result = await sendCommand('executeCode', { code });
      return {
        content: [{ type: "text", text: JSON.stringify(result ?? { success: true }, null, 2) }],
      };
    } catch (err) {
      return { content: [{ type: "text", text: (err as Error).message }], isError: true };
    }
  }
);

// Start MCP transport AFTER HTTPS server is listening
const transport = new StdioServerTransport();
await mcpServer.connect(transport);
console.error("MCP server connected via stdio");
```

### Tool Return Format Reference

```typescript
// Source: SDK type definitions (CallToolResult in types.d.ts)

// Success (text result):
return {
  content: [{ type: "text", text: "some result" }],
};

// Success (JSON result):
return {
  content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
};

// Error:
return {
  content: [{ type: "text", text: "Error: Add-in not connected" }],
  isError: true,
};
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `server.tool()` | `server.registerTool()` | SDK v1.20+ | `tool()` deprecated, `registerTool()` uses config object pattern |
| `zod@3.22` | `zod@^3.25` | SDK v1.20+ | SDK requires ^3.25 for internal Zod v4 compatibility layer |
| SSE transport | Streamable HTTP transport | SDK v1.18+ | SSE deprecated for new servers (not relevant here -- we use stdio) |

**Deprecated/outdated:**
- `server.tool()`: Deprecated in favor of `server.registerTool()`. Still works but will be removed in v2.
- SSE transport (`SSEServerTransport`): Deprecated in favor of `StreamableHTTP`. Not relevant since we use stdio.

## Node.js 24 Native TypeScript Compatibility

**Confidence: HIGH** (verified against SDK source)

The `@modelcontextprotocol/sdk` package ships pre-compiled JavaScript in `dist/esm/` and `dist/cjs/`. Node.js 24 imports the compiled `.js` files, not TypeScript source. Therefore:

- `erasableSyntaxOnly` has no effect on SDK code (it's already JS)
- The SDK works with Node.js >= 18 (verified in package.json `engines`)
- The project's own code must avoid TypeScript enums (use `as const` objects instead, which is already the pattern in the existing codebase)
- `verbatimModuleSyntax` requires using `import type` for type-only imports from zod/SDK if needed

The only user-side code that must be TypeScript-compatible is the tool registration in `server/index.ts`, which uses standard constructs (imports, async functions, arrow functions, type annotations) -- all compatible with `erasableSyntaxOnly`.

## Open Questions

1. **Office.js shape loading patterns**
   - What we know: Office.js uses a proxy-object model that requires `.load()` and `context.sync()` for each property batch.
   - What's unclear: The exact load pattern needed for `get_slide` (whether textFrame/fill properties can be loaded in a single batch or require sequential syncs). The code examples above use a conservative multi-sync approach.
   - Recommendation: Test the Office.js code strings against a live presentation during implementation. The code may need refinement based on which properties are available on different shape types.

2. **Shape type values**
   - What we know: Office.js `Shape.type` returns values from `ShapeType` enum (e.g., "GeometricShape", "Image", "Table", etc.).
   - What's unclear: The exact enum values available in PowerPoint's requirement set 1.1-1.9.
   - Recommendation: The executor will discover and report actual values. Not a blocker for tool registration.

## Sources

### Primary (HIGH confidence)
- `@modelcontextprotocol/sdk` v1.26.0 package inspection -- mcp.d.ts type definitions, package.json exports/engines
- Official MCP Quickstart Guide (modelcontextprotocol.io/docs/develop/build-server) -- complete TypeScript example
- SDK npm registry metadata -- version 1.26.0, peer dependencies, engines

### Secondary (MEDIUM confidence)
- MCP Server Building Guide (mcpcat.io/guides/building-stdio-mcp-server) -- stdio gotchas and console.log warning
- Official MCP GitHub README (github.com/modelcontextprotocol/typescript-sdk) -- v1.x vs v2 status

### Tertiary (LOW confidence)
- Blog posts on MCP tool registration patterns -- confirmed against SDK type definitions

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- verified against npm registry, SDK source, and official docs
- Architecture: HIGH -- verified that stdio + HTTPS coexist (independent I/O channels), confirmed API signatures from type definitions
- Pitfalls: HIGH -- console.log corruption is well-documented across multiple sources and official docs
- Office.js code patterns: MEDIUM -- based on existing working code in codebase (the /api/test endpoint works), but complex load patterns need live testing

**Research date:** 2026-02-08
**Valid until:** 2026-03-08 (stable -- MCP SDK v1.x is in maintenance mode)
