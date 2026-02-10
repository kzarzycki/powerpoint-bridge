# PowerPoint Office.js Bridge - Research Findings

## Goal

Build a custom Office.js add-in + WebSocket bridge + MCP server that lets Claude Code manipulate live PowerPoint presentations on macOS. No existing solution does this - all macOS MCP servers use python-pptx (file-based only).

---

## 1. Office.js PowerPoint API Capabilities

### Requirement Sets (macOS Support: 1.1-1.9 stable, 1.10 preview)

| Set | Capabilities | macOS Version |
|-----|-------------|---------------|
| 1.1 | Create new presentations | 16.19+ |
| 1.2 | Insert/delete slides from other presentations (Base64) | 16.19+ |
| 1.3 | Add/delete slides; custom metadata tags | 16.19+ |
| 1.4 | Shape manipulation (add, move, size, format, delete geometric shapes, lines, textboxes) | 16.19+ |
| 1.5 | Hyperlink management; select slides, text ranges, shapes | 16.19+ |
| 1.6 | Select slides, text ranges, shapes programmatically | 16.19+ |
| 1.7 | Custom and document properties | 16.19+ |
| 1.8 | Shape bindings, tables | 16.19+ |
| 1.9 | Table formatting and management | 16.19+ |
| 1.10 | Accessibility, slide backgrounds, hyperlinks (PREVIEW) | 16.105+ |

### Key API Patterns

```javascript
// Add geometric shape
const shapes = slide.shapes;
const rect = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
rect.left = 100; rect.top = 100;
rect.height = 150; rect.width = 150;

// Set text on shape
shape.textFrame.textRange.text = "Shape text";
shape.textFrame.textRange.font.color = "purple";
shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;

// Fill color (solid only)
shape.fill.setSolidColor("lightblue");

// Add text box
shapes.addTextBox("Hello World");

// Add line
shapes.addLine(PowerPoint.ConnectorType.straight, {left: 0, top: 0, height: 100, width: 100});

// Group shapes
const group = shapes.addGroup(shapesToGroup);
group.group.ungroup(); // to ungroup

// Bindings (persistent shape references)
Office.context.document.bindings.add(shape, bindingType, {id: 'uniqueId'});

// Delete all shapes
shapes.items.forEach(shape => shape.delete());
```

### Critical Limitations

- **NO image insertion API** - only workaround is Base64 slide import via `presentation.insertSlidesFromBase64()`
- **NO chart creation** through JavaScript API
- **NO animation/transition control** in stable APIs
- **NO slide master editing**
- **NO theme/master layout application** to existing slides
- **Limited formatting** - solid fills and font color only, no gradients/effects/shadows
- Positioning uses **points** (1/72 inch): `left`, `top`, `height`, `width`

---

## 2. Add-in Architecture on macOS

### Runtime Environment

- Runs inside **Safari WKWebView** (WebKit2), not Edge WebView2 like Windows
- Lives in browser sandbox - **cannot host servers or bind to ports**
- **CAN make outbound connections**: WebSocket client, HTTP/Fetch, XMLHttpRequest
- Full WebSocket API support confirmed
- Modern ES6+ JavaScript supported
- No "windowless hosting" on macOS

### Add-in Components

1. **Manifest file** - XML defining URL endpoints, integration points, permissions
2. **Web application** - HTML/CSS/JS served from local or hosted URL
3. **Office.js library** - bridge for host application communication
4. **Taskpane** - UI panel in PowerPoint sidebar

### Security Constraints

- WKWebView enforces HTTPS/WSS for network connections
- Self-signed certificates need macOS Keychain trust setup
- Chrome 142+ blocks local network access for Office Online (Issue #6281) - desktop Office unaffected
- SecurityError reported with WebSockets in some configurations (Issue #1018)

### Sideloading on macOS

Add-in manifests go in: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`

---

## 3. WebSocket Bridge Pattern (Primary Architecture)

```
Claude Code CLI
    | (MCP stdio protocol)
    v
MCP Server (Node.js, localhost)
    | (WebSocket server on wss://localhost:PORT)
    v
WebSocket Bridge Server
    ^ (WS client connects on add-in load)
    |
PowerPoint Add-in (inside WKWebView)
    | (Office.js API calls)
    v
PowerPoint Document (live, open)
```

### Data Flow

1. Claude Code sends MCP tool calls (e.g., `add_shape`, `set_text`)
2. MCP server translates to JSON command messages
3. WebSocket server pushes command to connected add-in client
4. Add-in executes Office.js API calls against live presentation
5. Results flow back through WebSocket → MCP → Claude Code

### Implementation Considerations

- **WSS required** - WKWebView won't connect to plain `ws://localhost`
- **Self-signed cert** - generate with `mkcert` for localhost, trust in macOS Keychain
- **Reconnection logic** - add-in must handle WS disconnects gracefully
- **Command/response protocol** - JSON with request IDs for async matching
- **Heartbeat/ping** - keep connection alive

### Alternative Patterns Considered

| Pattern | Pros | Cons |
|---------|------|------|
| WebSocket (chosen) | Real-time, bidirectional | WSS cert setup |
| HTTP polling | Simpler, no cert issues | Higher latency, wasteful |
| SSE (Server-Sent Events) | One-way push | No response channel |

---

## 4. Existing Solutions Landscape

### MCP Servers for PowerPoint

| Project | Tech | macOS | Live Edit | Tools |
|---------|------|-------|-----------|-------|
| socamalo/PPT_MCP_Server | COM (pywin32) | NO | Yes (Win) | 12 |
| GongRzhe/Office-PowerPoint-MCP-Server | python-pptx | YES | NO | 34 |
| vancealexander/Powerpoint_MCP_CrossPlatform | COM/python-pptx | YES | NO (macOS) | 13 |
| SlideSpeak MCP | Cloud API | YES | NO | - |
| jenstangen1/pptx-xlsx-mcp | python-pptx | YES | NO | - |

**Key insight:** ALL macOS solutions use python-pptx (file-based). NONE can edit open presentations. Only Windows COM solutions have live editing. Our Office.js bridge would be the first live-editing solution for macOS.

### Anthropic's Built-in PPTX Skill

- Uses **PptxGenJS** (Node.js) for creation + **markitdown** for reading
- Generates .pptx in VM sandbox, converts to PDF/images for QA
- Source: `github.com/anthropics/skills/blob/main/skills/pptx/SKILL.md`
- Typography rules: 36-44pt titles, 20-24pt headers, 14-16pt body
- 10 color palettes (Midnight Executive, Forest & Moss, etc.)
- QA loop: content check → visual inspection → verification
- **Limitation:** file generation only, no live editing

### Other Related Projects (Windows-only)

- powerpoint-remote-websocket (VSTO, slideshow control)
- PowerStage (WebSocket remote for presentations)
- OSCPoint (OSC protocol control)

---

## 5. macOS-Specific Considerations

### WKWebView Runtime

- WebKit2 engine, modern standards support
- Full WebSocket client API available
- CSP (Content Security Policy) must allow WSS connections
- `manifest.xml` AppDomains must include localhost

### Certificate Setup for WSS

```bash
# Install mkcert
brew install mkcert
mkcert -install  # Adds local CA to macOS Keychain

# Generate localhost cert
mkcert localhost 127.0.0.1 ::1
# Produces: localhost+2.pem and localhost+2-key.pem
```

### Known macOS Issues

- PowerPoint Mac 16.94 (Feb 2025): `ActiveViewChanged` event stopped firing
- Previous version 16.93.2 (Jan 2025) worked fine
- No Conditional Access enterprise features on macOS

### Sideloading Path

```
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```
Place `manifest.xml` here for development sideloading.

---

## 6. References

### Microsoft Learn Documentation

- [PowerPoint JS API Reference](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/powerpoint-add-ins-reference-overview)
- [Requirement Sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)
- [Working with Shapes](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/shapes)
- [Browsers Used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)
- [Script Lab](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/explore-with-script-lab)

### GitHub Issues (Relevant)

- [#236 - Cannot host server in add-in](https://github.com/OfficeDev/office-js-docs/issues/236)
- [#70 - WebSocket support](https://github.com/OfficeDev/office-js-docs/issues/70)
- [#1018 - SecurityError with WebSockets](https://github.com/OfficeDev/office-js/issues/1018)
- [#6281 - Chrome 142+ local network access block](https://github.com/OfficeDev/office-js/issues/6281)

### Existing MCP Servers

- [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) - most comprehensive python-pptx
- [socamalo/PPT_MCP_Server](https://github.com/socamalo/PPT_MCP_Server) - Windows COM live editing
- [vancealexander/Powerpoint_MCP_CrossPlatform](https://glama.ai/mcp/servers/@vancealexander/Powerpoint_MCP_CrossPlatform)

### Other References

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Anthropic Skills - pptx](https://github.com/anthropics/skills/blob/main/skills/pptx/SKILL.md)
- [MacTech AppleScript PowerPoint Guide](http://preserve.mactech.com/articles/mactech/Vol.23/23.03/23.03AppleScriptPowerPoint/index.html)
