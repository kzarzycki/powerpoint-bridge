---
name: powerpoint-live
description: "Manipulate live, open PowerPoint presentations on macOS via Office.js MCP bridge. Use when Claude needs to: (1) create, edit, or inspect slides in a running PowerPoint instance, (2) add shapes, text, tables, or formatting to live presentations, (3) capture visual slide screenshots, (4) enable/configure the PowerPoint MCP bridge in a project, (5) execute Office.js code against open presentations. Distinct from the pptx file-editing skill — this works on presentations currently open in PowerPoint."
---

# PowerPoint Live Editing

Edit live, open PowerPoint presentations through an MCP bridge. Changes appear in real-time.

```
Claude Code  ──MCP HTTP (localhost:3001)──>  Bridge Server  ──WSS──>  PowerPoint Add-in  ──>  Live Presentation
```

## Enabling in a Project

When asked to enable, configure, or set up PowerPoint MCP in a project:

1. Add to the project's `.mcp.json` (create if missing):
   ```json
   {
     "mcpServers": {
       "powerpoint-bridge": {
         "type": "http",
         "url": "http://localhost:3001/mcp"
       }
     }
   }
   ```
2. Verify server is reachable: `curl -s http://localhost:3001/mcp -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'`
3. If server unreachable, direct user: "Start the bridge server: `cd ~/powerpoint-bridge && npm start`"
4. Test: call `list_presentations` to confirm connectivity

If the bridge has never been installed, see [setup guide](references/setup.md).

## MCP Tools

| Tool | Purpose | Key Parameters |
|------|---------|---------------|
| `list_presentations` | Discover connected presentations | — |
| `get_presentation` | Get slide structure (indices, shape names/types) | `presentationId?` |
| `get_slide` | Get detailed shapes (positions, text, colors) | `slideIndex`, `presentationId?` |
| `get_slide_image` | Capture slide screenshot as PNG | `slideIndex`, `width?` (default 720), `presentationId?` |
| `export_slide` | Export slide as Base64 .pptx (for copying between presentations) | `slideIndex`, `presentationId?` |
| `insert_slides` | Insert slides from Base64 .pptx into current presentation | `base64`, `formatting?`, `targetSlideId?`, `sourceSlideIds?`, `presentationId?` |
| `execute_officejs` | Run arbitrary Office.js code in the live presentation | `code`, `presentationId?` |

`presentationId` is required only when multiple presentations are connected. Get it from `list_presentations`.

All positioning values from `get_slide` are in **points** (1 pt = 1/72 inch). Standard 16:9 slide: 960 x 540 pt.

## Workflow

1. **Discover**: `list_presentations` — find connected presentations
2. **Understand**: `get_presentation` then `get_slide` — learn structure
3. **See**: `get_slide_image` — visually inspect current state
4. **Modify**: `execute_officejs` — make changes with Office.js code
5. **Verify**: `get_slide_image` — confirm visual result

Always inspect before modifying. Always verify after modifying.

For `execute_officejs` code patterns, see [code-patterns.md](references/code-patterns.md).

## Hard Limitations

Cannot do via Office.js — do not attempt:

- Insert images directly (workaround: `presentation.insertSlidesFromBase64()`)
- Create or edit charts
- Add animations or transitions
- Apply gradients, shadows, or effects (solid fills only)
- Edit slide masters or themes

## Error Handling

- **"No presentations connected"** — open PowerPoint with the add-in loaded
- **"Multiple presentations connected"** — specify `presentationId`
- **"Add-in disconnected"** — auto-reconnects; wait and retry
- **"Command timed out"** — simplify code or check PowerPoint responsiveness
- **Screenshot via execute_officejs overflows tokens** — always use `get_slide_image` instead (returns MCP image block, not text)
