# PowerPoint Office.js Bridge

## Project Goal

Build a system that lets Claude Code manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs. First live-editing MCP solution for macOS — all others use python-pptx (file-based).

## Architecture

```
Claude Code  <--MCP STDIO/HTTP-->  Bridge Server (Node.js)  <--WS-->  PowerPoint Add-in (Office.js)
                                          |                                     |
                                    STDIO (default)                      WKWebView sandbox
                                    or HTTP (:3001/mcp)                  Office.js API 1.1-1.9
                                    localhost:8080 (HTTP)                executes commands on
                                    serves add-in files + WS             live presentation
```

Three components in one repo:

### 1. `addin/` - Office.js PowerPoint Add-in
- HTML/CSS/JS taskpane that loads inside PowerPoint
- Connects as WebSocket **client** to bridge server on load
- Receives JSON commands, executes Office.js API calls, returns results
- Manifest XML for sideloading

### 2. `server/` - Bridge Server (Node.js)
- HTTP server serving add-in static files (HTTPS opt-in via `BRIDGE_TLS=1`)
- WS (WebSocket) server for add-in connection (WSS when TLS enabled)
- MCP server (STDIO default, HTTP on port 3001 for standalone) exposing tools to Claude Code
- All three roles in one process for simplicity

### 3. `certs/` - Optional Local TLS Certificates
- Generated via `mkcert` for localhost (only needed when `BRIDGE_TLS=1`)
- `.gitignore`d

## Usage Documentation

For tool reference, code patterns, and usage — see the **powerpoint-live** skill at `skills/powerpoint-live/`.

## Technical Constraints

- **HTTPS/WSS opt-in** - Set `BRIDGE_TLS=1` to use HTTPS/WSS (requires certs via `npm run setup-certs`)
- **Add-in cannot host servers** - sandboxed in WKWebView, can only make outbound connections
- **Limited image API** - Image insertion via Common API `setSelectedDataAsync` (`insert_image` tool); no shape-level `addPicture()` yet (BETA only)
- **No charts** - Office.js cannot create charts
- **No animations** - not exposed in stable APIs
- **Solid fills only** - no gradients, effects, or shadows
- **Points for positioning** - 1 point = 1/72 inch

## Conventions

- PR titles and commits follow Conventional Commits (`feat:`, `fix:`, `chore:`, etc.)

## Development Workflow

### Build
- **After changing `server/` files**: Run `npm run build` and stage `dist/index.cjs` — it's committed and used by plugin/MCPB installs. The pre-commit hook enforces this.
- **Quality gate**: Use `npm run check` (runs lint + typecheck + test in one command)

### Branch & PR
- Create a feature branch: `<type>/<short-description>` (e.g. `fix/tcc-sideload-prompt`, `feat/new-tool`)
- PR title must follow Conventional Commits — CI enforces this and auto-labels (`feat`→enhancement, `fix`→bug, etc.)
- After pushing, always wait for CI checks to pass before confirming the PR is ready
- Squash merge PRs to keep main history clean
- Delete feature branch after merge

### Code Quality
- **TDD (red-green)**: Write failing tests first, then implement to make them pass
- After implementation, use `/simplify` to review changed code
- Run `npm run check` before pushing

### GitHub Actions (CI Claude)

When running in GitHub Actions via `@claude` mentions, follow this workflow model:

**Issues = requirements level (what/why)**
- `@claude research this` → Analyze the problem at user/requirements level: what's broken, who's affected, what should work differently, acceptance criteria. Do NOT include file paths, code snippets, or implementation details.
- `@claude implement this` (small/obvious fix) → Go straight to creating a PR with fix.
- `@claude implement this` (medium+ task) → Create a PR with technical plan in description, then implement. The PR description becomes the technical design doc.

**PRs = solution level (how)**
- Technical research, planning, and implementation happen here.
- PR description should contain: technical approach, files to modify, trade-offs considered.
- `@claude plan the approach` → Post technical plan as PR comment.
- `@claude implement this` → Write code, run `npm run check`, push.
- `@claude fix X` → Iterate on implementation.
- Always run `npm run check` before marking work as done.
- PR title must follow Conventional Commits.

**Deciding task size:**
- Trivial (typo, config, one-liner): skip planning, go straight to PR with fix.
- Small (clear bug, isolated change): create PR with brief plan in description + code.
- Medium (feature, multi-file refactor): create PR with detailed plan in description + code.
- If uncertain about scope or approach, ask in a comment instead of guessing.

## Key Technical Decisions

1. **Single Node.js process** for HTTP(S) + WS(S) + MCP (simplicity over microservices)
2. **TypeScript** for add-in and server (Office.js has good TS types)
3. **JSON command protocol** with request IDs for async response matching
4. **Plain HTTP/WS by default**, HTTPS/WSS opt-in via `BRIDGE_TLS=1` + mkcert certs
5. **Sideloading** for development (no Microsoft store submission needed)

## Command Protocol (WebSocket Messages)

```typescript
// Client (add-in) → Server
interface WSMessage {
  type: 'response' | 'error' | 'ready';
  id?: string;        // matches request ID
  data?: any;
}

// Server → Client (add-in)
interface WSCommand {
  type: 'command';
  id: string;         // unique request ID
  action: string;     // e.g. 'addShape', 'setText', 'getSlides'
  params: Record<string, any>;
}
```

## References

See `RESEARCH.md` for full research findings including:
- Detailed API capabilities per requirement set
- Code examples for all shape/text/table operations
- Existing solutions comparison
- macOS-specific issues and workarounds
- All relevant Microsoft documentation and GitHub issue links
