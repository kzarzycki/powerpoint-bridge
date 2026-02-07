# Phase 3: Command Execution - Research

**Researched:** 2026-02-07
**Domain:** WebSocket JSON protocol design, dynamic Office.js code execution, request/response matching
**Confidence:** HIGH

## Summary

Phase 3 wires up the command execution pipeline: the server sends a JSON command with a unique request ID over WebSocket, the add-in receives it, executes Office.js code inside `PowerPoint.run()`, and returns the result (or error) as a JSON response with the matching ID. This is the core plumbing that Phase 4's MCP tools will use.

The key technical insight is that Office.js add-ins on macOS do NOT enforce a Content Security Policy that blocks `eval()` or `new Function()`. Office.js itself requires `unsafe-eval` (via MicrosoftAjax.js injection). Since we control the HTTPS server and send no restrictive CSP headers, the `AsyncFunction` constructor works in the add-in's WKWebView. This enables the `execute_officejs` pattern: send a code string, wrap it in an `AsyncFunction`, call it inside `PowerPoint.run()`, and return the result. `PowerPoint.run()` propagates the callback's return value through its Promise (it returns `Promise<T>`, not `Promise<void>`), so code strings can `return` serializable values that flow back to the server.

The server side uses a `Map<string, {resolve, reject}>` keyed by request ID to match incoming responses to pending commands. `crypto.randomUUID()` (built into Node 24) generates request IDs. A timeout per request prevents hanging if the add-in never responds.

**Primary recommendation:** Build a single `executeCode` action that wraps code strings in an `AsyncFunction('context', 'PowerPoint', code)` and executes them inside `PowerPoint.run()`. This is the only action type needed -- Phase 4's `get_presentation`, `get_slide`, and `execute_officejs` MCP tools all send code strings through this same mechanism.

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| `crypto.randomUUID()` | Node 24 built-in | Generate unique request IDs | RFC 4122 v4 UUID, no external dependency, available since Node 19 |
| `AsyncFunction` constructor | JS built-in | Create async functions from code strings | Enables `await` inside dynamic code; obtained via `(async function(){}).constructor` |
| `ws` (existing) | ^8.19.0 | WebSocket server | Already installed and working from Phase 1 |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| No new dependencies | - | - | Phase 3 requires zero new npm packages |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `AsyncFunction` constructor | `eval()` with async wrapper | Both work; AsyncFunction is cleaner, takes named args, returns a function object |
| `crypto.randomUUID()` | `uuid` npm package | Built-in is faster and zero-dependency; uuid package is unnecessary on Node 19+ |
| Single `executeCode` action | Multiple discrete actions (addShape, getSlides, etc.) | Discrete actions would need a handler per action in the add-in; `executeCode` is maximally flexible and aligns with the 3-tool MCP architecture decided in PROJECT.md |
| Custom protocol | JSON-RPC 2.0 | JSON-RPC adds batch request and notification concepts we don't need; simpler custom protocol is sufficient for single-client WebSocket |

**Installation:**
```bash
# No new packages needed
```

## Architecture Patterns

### Recommended Changes to Existing Files
```
server/index.ts      # Add: pending request tracking, sendCommand(), message handler dispatch
addin/app.js         # Add: handleCommand() implementation, executeCode(), sendResponse(), sendError()
```

No new files needed. Both changes extend existing files from Phases 1-2.

### Pattern 1: Pending Request Map (Server Side)
**What:** A Map that stores Promise resolve/reject callbacks keyed by request ID. When the add-in responds, the matching callback is invoked.
**When to use:** Every command sent to the add-in.
**Example:**
```typescript
// Source: Standard WebSocket request-response correlation pattern
import { randomUUID } from 'node:crypto';

const pendingRequests = new Map<string, {
  resolve: (data: unknown) => void;
  reject: (err: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}>();

const COMMAND_TIMEOUT = 30_000; // 30 seconds

function sendCommand(ws: WebSocket, action: string, params: Record<string, unknown>): Promise<unknown> {
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      reject(new Error(`Command ${id} timed out after ${COMMAND_TIMEOUT}ms`));
    }, COMMAND_TIMEOUT);

    pendingRequests.set(id, { resolve, reject, timer });
    ws.send(JSON.stringify({ type: 'command', id, action, params }));
  });
}
```

### Pattern 2: Response Dispatch (Server Side)
**What:** When a WebSocket message arrives from the add-in, look up the pending request by ID and resolve/reject it.
**When to use:** In the `ws.on('message')` handler.
**Example:**
```typescript
ws.on('message', (raw: Buffer) => {
  const msg = JSON.parse(raw.toString());

  if ((msg.type === 'response' || msg.type === 'error') && msg.id) {
    const pending = pendingRequests.get(msg.id);
    if (pending) {
      clearTimeout(pending.timer);
      pendingRequests.delete(msg.id);
      if (msg.type === 'response') {
        pending.resolve(msg.data);
      } else {
        const err = new Error(msg.error?.message || 'Command failed');
        pending.reject(err);
      }
    }
  }

  if (msg.type === 'ready') {
    console.log('Add-in ready to receive commands');
  }
});
```

### Pattern 3: AsyncFunction Code Execution (Add-in Side)
**What:** Create an async function from a code string and execute it inside `PowerPoint.run()`. The code string receives `context` (the RequestContext) and `PowerPoint` (the namespace for enums/constants) as arguments.
**When to use:** For every `executeCode` action.
**Example:**
```javascript
// Source: MDN AsyncFunction constructor + Office.js PowerPoint.run()
var AsyncFunction = (async function(){}).constructor;

function handleCommand(message) {
  if (message.type !== 'command' || !message.id) return;

  if (message.action === 'executeCode') {
    executeCode(message.params.code, message.id);
  } else {
    sendError(message.id, { message: 'Unknown action: ' + message.action });
  }
}

function executeCode(code, requestId) {
  PowerPoint.run(async function(context) {
    var fn = new AsyncFunction('context', 'PowerPoint', code);
    var result = await fn(context, PowerPoint);
    return result;
  }).then(function(result) {
    // Convert undefined to null for JSON serialization
    sendResponse(requestId, result === undefined ? null : result);
  }).catch(function(error) {
    sendError(requestId, {
      message: error.message || String(error),
      code: error.code || 'UnknownError',
      debugInfo: error.debugInfo || null
    });
  });
}
```

### Pattern 4: Ready Signal (Add-in Side)
**What:** After WebSocket connection opens and Office.js is ready, send a `{type: 'ready'}` message to tell the server the add-in can receive commands.
**When to use:** On WebSocket open, after Office.js has initialized.
**Example:**
```javascript
ws.onopen = function() {
  reconnectAttempt = 0;
  updateStatus('connected');
  // Signal readiness to receive commands
  ws.send(JSON.stringify({ type: 'ready' }));
};
```

### Pattern 5: PowerPoint.run() Return Value Propagation
**What:** `PowerPoint.run<T>(callback)` returns `Promise<T>` -- the callback's return value propagates through the Promise chain. Code strings can `return` values that become the resolved result.
**When to use:** Always -- this is how results flow back from Office.js execution.
**Example:**
```javascript
// Code string sent from server:
// "var count = context.presentation.slides.getCount(); await context.sync(); return count.value;"
//
// Inside executeCode, this becomes:
// PowerPoint.run(async function(context) {
//   var fn = new AsyncFunction('context', 'PowerPoint', codeString);
//   var result = await fn(context, PowerPoint);  // result = 3 (or whatever the count is)
//   return result;
// }).then(function(result) {
//   sendResponse(requestId, result);  // sends 3 back to server
// });
```

### Anti-Patterns to Avoid
- **Returning Office.js proxy objects from code strings:** Proxy objects are not JSON-serializable. Code must extract primitive values (strings, numbers, arrays, plain objects) before returning. Always `load()` + `context.sync()` + read `.value` or properties, then return plain data.
- **Forgetting `await context.sync()` before reading loaded properties:** The load/sync pattern is mandatory. Without sync, loaded properties return undefined.
- **Sending commands before 'ready' signal:** The add-in may connect before Office.js initializes. Sending commands too early will fail because `PowerPoint.run()` is not available.
- **Not handling WebSocket disconnection in pending requests:** If the add-in disconnects, all pending requests should be rejected, not left hanging forever.
- **Using `eval()` instead of `AsyncFunction`:** While `eval()` works, it executes in the current scope and doesn't support `await` at the top level. `AsyncFunction` constructor creates a proper async function with named parameters.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| UUID generation | Custom counter or random string | `crypto.randomUUID()` | Built-in, cryptographically secure, RFC-compliant, no dependencies |
| Async dynamic code execution | `eval()` with Promise wrapper | `AsyncFunction` constructor | Clean parameter passing, proper async/await support, function object semantics |
| Multiple discrete action handlers | A handler per Office.js operation (addShape, setText, getSlides...) | Single `executeCode` action | The 3-tool MCP architecture (from PROJECT.md) means Claude writes Office.js code directly. The add-in only needs one action type. Discrete handlers would be duplicated effort that goes unused. |
| JSON parsing/serialization | Custom protocol encoding | `JSON.parse()` / `JSON.stringify()` | Standard, debuggable, matches the WebSocket text frame format already in use |

**Key insight:** Phase 3 is a thin execution layer. The intelligence lives in the MCP tools (Phase 4) that generate the code strings. The add-in just needs to run code and return results. One action type (`executeCode`) is sufficient.

## Common Pitfalls

### Pitfall 1: Non-Serializable Return Values
**What goes wrong:** Code string returns an Office.js proxy object (e.g., a Shape or Slide object). `JSON.stringify()` either throws or produces `{}` because proxy objects have no own enumerable properties until loaded.
**Why it happens:** Office.js uses a proxy pattern -- objects don't hold data until `load()` + `sync()`. Returning the proxy instead of extracted values sends nothing useful back.
**How to avoid:** Code strings must always extract primitive values before returning. Pattern: `load('properties') -> sync() -> read .property -> return plain object`.
**Warning signs:** Response `data` is `null`, `{}`, or `undefined` when expecting structured data.

### Pitfall 2: Missing context.sync() Before Reading Properties
**What goes wrong:** Code loads properties but reads them before calling `await context.sync()`, getting `undefined` for every property.
**Why it happens:** `load()` only queues a read request. Values aren't populated until `sync()` executes the queued operations.
**How to avoid:** Always follow the pattern: `obj.load('props'); await context.sync(); return obj.props;`
**Warning signs:** All property values are `undefined` even though the object exists.

### Pitfall 3: Command Sent Before Add-in Ready
**What goes wrong:** Server sends a command immediately after WebSocket connection, but `PowerPoint.run()` is not yet available in the add-in.
**Why it happens:** WebSocket connects before `Office.onReady()` fires (or before the 3-second fallback in standalone mode).
**How to avoid:** Add-in sends `{type: 'ready'}` only after both Office.js initialization and WebSocket connection succeed. Server waits for this signal before sending commands.
**Warning signs:** Error response with "PowerPoint is not defined" or "Office is not defined".

### Pitfall 4: Pending Requests Leak on Disconnect
**What goes wrong:** Add-in disconnects while commands are pending. The Promises never resolve or reject, causing memory leaks and hanging callers.
**Why it happens:** The response dispatch only runs when a message arrives. No message arrives if the WebSocket closes.
**How to avoid:** On WebSocket close, iterate `pendingRequests` and reject all with a "disconnected" error, then clear the map.
**Warning signs:** Commands hang indefinitely; server memory grows over time.

### Pitfall 5: AsyncFunction Not Available as Global
**What goes wrong:** Code tries `new AsyncFunction(...)` but gets "AsyncFunction is not defined" because it is not a global constructor.
**Why it happens:** Unlike `Function`, `AsyncFunction` is not exposed as a global identifier in JavaScript.
**How to avoid:** Obtain it via: `var AsyncFunction = (async function(){}).constructor;`
**Warning signs:** ReferenceError on first command execution.

### Pitfall 6: Error Object Serialization
**What goes wrong:** Catching an Office.js error and sending `JSON.stringify(error)` produces `{}` because Error objects have non-enumerable properties (`message`, `stack`).
**Why it happens:** `JSON.stringify()` only serializes own enumerable properties. `Error.prototype.message` and `Error.prototype.stack` are not enumerable.
**How to avoid:** Manually extract `error.message`, `error.code`, and `error.debugInfo` into a plain object before sending.
**Warning signs:** Error responses have empty `error` objects.

## Code Examples

### Complete Server-Side Command Infrastructure
```typescript
// Source: Verified Office.js PowerPoint.run() signature + standard WebSocket request-response pattern
import { randomUUID } from 'node:crypto';
import type { WebSocket } from 'ws';

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

// In wss.on('connection') handler:
// - Set addinClient = ws
// - Parse incoming messages and dispatch to pendingRequests
// - On close: reject all pending, set addinClient = null, addinReady = false
```

### Complete Add-in Command Handler
```javascript
// Source: MDN AsyncFunction constructor + Office.js PowerPoint.run() + error handling docs
var AsyncFunction = (async function(){}).constructor;

function handleCommand(message) {
  if (message.type !== 'command' || !message.id) return;

  if (message.action === 'executeCode') {
    executeCode(message.params.code, message.id);
  } else {
    sendError(message.id, { message: 'Unknown action: ' + message.action });
  }
}

function executeCode(code, requestId) {
  PowerPoint.run(async function(context) {
    var fn = new AsyncFunction('context', 'PowerPoint', code);
    var result = await fn(context, PowerPoint);
    return result;
  }).then(function(result) {
    sendResponse(requestId, result === undefined ? null : result);
  }).catch(function(error) {
    var errorObj = {
      message: error.message || String(error),
      code: error.code || 'UnknownError'
    };
    if (error.debugInfo) {
      errorObj.debugInfo = error.debugInfo;
    }
    sendError(requestId, errorObj);
  });
}

function sendResponse(id, data) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'response', id: id, data: data }));
  }
}

function sendError(id, error) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'error', id: id, error: error }));
  }
}
```

### Test Code String: Get Slide Count
```javascript
// This is the code string the server sends to verify the bridge works
// It runs inside PowerPoint.run() via AsyncFunction
var count = context.presentation.slides.getCount();
await context.sync();
return count.value;
```

### Test Code String: Get Slide IDs
```javascript
// Returns array of slide IDs from the open presentation
context.presentation.slides.load('items/id');
await context.sync();
var ids = [];
for (var i = 0; i < context.presentation.slides.items.length; i++) {
  ids.push(context.presentation.slides.items[i].id);
}
return ids;
```

### Test Code String: Add a Shape
```javascript
// Adds a blue rectangle to the first slide and returns its name
var slide = context.presentation.slides.getItemAt(0);
var shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
  left: 100, top: 100, width: 200, height: 100
});
shape.name = 'TestRect';
shape.fill.setSolidColor('#4472C4');
shape.textFrame.textRange.text = 'Hello from Bridge';
shape.textFrame.textRange.font.color = '#FFFFFF';
await context.sync();
return shape.name;
```

## JSON Protocol Specification

### Server to Add-in (Command)
```json
{
  "type": "command",
  "id": "550e8400-e29b-41d4-a716-446655440000",
  "action": "executeCode",
  "params": {
    "code": "var count = context.presentation.slides.getCount(); await context.sync(); return count.value;"
  }
}
```

### Add-in to Server (Success Response)
```json
{
  "type": "response",
  "id": "550e8400-e29b-41d4-a716-446655440000",
  "data": 3
}
```

### Add-in to Server (Error Response)
```json
{
  "type": "error",
  "id": "550e8400-e29b-41d4-a716-446655440000",
  "error": {
    "message": "The requested resource doesn't exist.",
    "code": "ItemNotFound",
    "debugInfo": null
  }
}
```

### Add-in to Server (Ready Signal)
```json
{
  "type": "ready"
}
```

## Office.js API Surface Relevant to Code Execution

### Key Objects and Their Loadable Properties

**Presentation** (via `context.presentation`):
- `title` (string), `id` (string, API 1.5)
- `slides` (SlideCollection)
- `getSelectedSlides()`, `getSelectedShapes()`

**Slide** (via `slides.getItemAt(index)` or `slides.getItem(id)`):
- `id` (string), `index` (number, API 1.8)
- `shapes` (ShapeCollection)
- `layout`, `slideMaster`, `tags`, `hyperlinks`
- Methods: `delete()`, `moveTo()`, `getImageAsBase64()`, `exportAsBase64()`

**Shape** (via `slide.shapes.getItemAt(index)`):
- `id`, `name`, `type` (ShapeType), `left`, `top`, `width`, `height` (all API 1.4)
- `fill` (ShapeFill): `foregroundColor`, `type`, `transparency`
- `lineFormat` (ShapeLineFormat): `color`, `weight`, `dashStyle`, `style`, `visible` (API 1.4)
- `textFrame` (TextFrame): `hasText`, `textRange`, `verticalAlignment`, `wordWrap` (API 1.4)
- `zOrderPosition`, `level`, `group`, `parentGroup` (API 1.8)
- `rotation`, `visible`, `altTextDescription`, `altTextTitle` (API 1.10)
- Methods: `delete()`, `getParentSlide()`, `setZOrder()`, `getTable()` (API 1.8)

**TextRange** (via `shape.textFrame.textRange`):
- `text` (string), `font` (ShapeFont), `paragraphFormat`, `start`, `length`
- Methods: `getSubstring()`, `setSelected()`

**ShapeFont** (via `textRange.font`):
- `bold`, `italic`, `color`, `name`, `size`, `underline` (all API 1.4)
- `allCaps`, `strikethrough`, `subscript`, `superscript` (API 1.8)

**ShapeCollection** (via `slide.shapes`):
- `items` (Shape[]), `getCount()`, `getItem(id)`, `getItemAt(index)`
- `addGeometricShape()`, `addLine()`, `addTextBox()`, `addTable()` (API 1.4-1.8)
- `addGroup()` (API 1.8)

### GeometricShapeType Values (Partial List)
`rectangle`, `roundRectangle`, `ellipse`, `triangle`, `diamond`, `pentagon`, `hexagon`, `octagon`, `star4`, `star5`, `star6`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `bracePair`, `bracketPair`, `cloud`, `heart`, `plus`

### ShapeType Values
`geometricShape`, `group`, `image`, `line`, `table`, `textBox`, `freeform`, `placeholder`, `unsupported`

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `uuid` npm package for IDs | `crypto.randomUUID()` built-in | Node.js 19 (Oct 2022) | No external dependency needed |
| `eval()` for dynamic code | `AsyncFunction` constructor | Available since ES2017 | Proper async/await support, named parameters, cleaner semantics |
| Discrete command handlers per action | Single code execution action | Architectural decision (PROJECT.md) | One action handles all operations; complexity moves to code generation |

**Deprecated/outdated:**
- `Office.initialize` callback: Still works but `Office.onReady()` is the modern approach (already using onReady in Phase 2)

## Open Questions

1. **PowerPoint.run() return value propagation in WKWebView**
   - What we know: The TypeScript signature is `run<T>(batch: (context) => IPromise<T>): IPromise<T>`, confirming the callback return value propagates. Multiple online examples use this pattern.
   - What's unclear: Whether WKWebView's specific JavaScript engine has any quirks with this. No reports of issues found.
   - Recommendation: Proceed with the return-value pattern. If it fails, the workaround is to store the result in a closure variable and read it after `PowerPoint.run()` resolves.

2. **Maximum code string size via WebSocket**
   - What we know: The `ws` library has a default `maxPayload` of 100 MB. WebSocket text frames have no practical size limit.
   - What's unclear: Whether very large code strings cause performance issues in WKWebView's JavaScript engine.
   - Recommendation: Not a concern for Phase 3. Code strings will be small (typical Office.js operations are 5-50 lines). If needed later, the ws library's `maxPayload` can be configured.

3. **Concurrent command execution**
   - What we know: Multiple commands could be sent before the first response returns. Each has a unique ID so responses can be matched.
   - What's unclear: Whether `PowerPoint.run()` can execute concurrently or if it serializes internally. The docs don't specify.
   - Recommendation: For Phase 3, commands will be sequential (one at a time). Phase 4's MCP tools are inherently sequential (one tool call at a time from Claude). If concurrency is needed later, add a queue.

## Sources

### Primary (HIGH confidence)
- [Microsoft Learn: PowerPoint.run() API reference](https://learn.microsoft.com/en-us/javascript/api/powerpoint) - Verified `run<T>` generic return type signature
- [Microsoft Learn: Application-specific API model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model) - Proxy object model, context.sync(), load() patterns, ClientResult<T>
- [Microsoft Learn: Working with Shapes](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/shapes) - Shape creation, text, fill, grouping code examples
- [Microsoft Learn: PowerPoint.Presentation class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation) - All properties/methods with requirement sets
- [Microsoft Learn: PowerPoint.Slide class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slide) - Slide properties, getImageAsBase64(), exportAsBase64()
- [Microsoft Learn: PowerPoint.Shape class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape) - All shape properties/methods with requirement sets
- [Microsoft Learn: PowerPoint.ShapeFill class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapefill) - foregroundColor, type, transparency, setSolidColor()
- [Microsoft Learn: PowerPoint.TextFrame class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textframe) - hasText, textRange, verticalAlignment
- [Microsoft Learn: PowerPoint.TextRange class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textrange) - text, font, getSubstring()
- [Microsoft Learn: PowerPoint.ShapeFont class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapefont) - bold, italic, color, name, size
- [Microsoft Learn: PowerPoint.ShapeLineFormat class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapelineformat) - color, weight, dashStyle
- [Microsoft Learn: PowerPoint.ShapeCollection class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection) - addGeometricShape, addLine, addTextBox, addTable
- [Microsoft Learn: PowerPoint.SlideCollection class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidecollection) - add, getCount, getItem, getItemAt
- [Microsoft Learn: Error handling in application-specific APIs](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/application-specific-api-error-handling) - Error codes, debugInfo, try-catch pattern
- [MDN: AsyncFunction() constructor](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/AsyncFunction/AsyncFunction) - Syntax, obtaining the constructor, security notes

### Secondary (MEDIUM confidence)
- [The Office Context: CSP and Office Add-ins](https://theofficecontext.com/2025/02/28/how-content-security-policy-affects-office-add-ins/) - Office.js injects MicrosoftAjax.js requiring unsafe-eval; no CSP enforced by Office runtime itself
- [Microsoft Learn: Privacy and security for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security) - Runtime sandbox model, process isolation, no mention of CSP enforcement
- [MDN: Crypto.randomUUID()](https://developer.mozilla.org/en-US/docs/Web/API/Crypto/randomUUID) - API reference for built-in UUID generation

### Tertiary (LOW confidence)
- [GitHub: rpc-websockets (JSON-RPC 2.0 over WebSocket)](https://github.com/elpheria/rpc-websockets) - Reviewed for protocol design patterns; decided simpler custom protocol is sufficient

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - All components are built-in (Node crypto, JS AsyncFunction, existing ws library). Zero new dependencies.
- Architecture: HIGH - PowerPoint.run() return type verified from official TypeScript signatures. AsyncFunction constructor verified from MDN. Protocol pattern is standard WebSocket request-response correlation.
- Pitfalls: HIGH - All pitfalls derived from verified Office.js documentation (proxy objects, load/sync pattern, error serialization) or standard JavaScript behavior (AsyncFunction not global, Error non-enumerable properties).
- CSP/eval compatibility: MEDIUM - Blog post confirms Office.js requires unsafe-eval. No official Microsoft docs explicitly state CSP policy for WKWebView. However, since we control the server and send no CSP header, there is no restriction to worry about.

**Research date:** 2026-02-07
**Valid until:** 2026-03-09 (30 days; Office.js API and Node.js built-ins are stable)
