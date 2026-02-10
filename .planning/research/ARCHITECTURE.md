# Architecture: Tooling Integration

**Domain:** How quality tooling integrates with the existing PowerPoint Bridge project
**Researched:** 2026-02-10

## Current Project Structure

```
powerpoint-bridge/
  addin/                  # Office.js add-in (HTML/CSS/JS)
  server/
    index.ts              # Single-file server (~700 LOC) - HTTPS + WSS + MCP
  certs/                  # TLS certs (gitignored)
  .planning/              # GSD planning files
  package.json
  tsconfig.json
  CLAUDE.md
  RESEARCH.md
```

## Target Structure After Tooling

```
powerpoint-bridge/
  .github/
    workflows/
      ci.yml              # GitHub Actions CI pipeline
  addin/                  # Office.js add-in (HTML/CSS/JS)
  server/
    index.ts              # Single-file server
  tests/
    tools.test.ts         # MCP tool handler tests
    protocol.test.ts      # Command protocol tests
    server.test.ts        # Server integration tests (if feasible)
  certs/                  # TLS certs (gitignored)
  biome.json              # Biome config (lint + format)
  vitest.config.ts        # Vitest config
  package.json            # Updated with scripts + metadata
  tsconfig.json           # No changes needed
  LICENSE                 # MIT license
  README.md               # Public-facing docs (replaces CLAUDE.md as primary)
  CLAUDE.md               # Agent instructions (kept for Claude Code users)
```

## Component Boundaries

| Component | Responsibility | Files |
|-----------|---------------|-------|
| Biome | Lint + format all TS/JS files | `biome.json` |
| Vitest | Run tests, generate coverage | `vitest.config.ts`, `tests/` |
| TypeScript | Type checking (no emit) | `tsconfig.json` |
| GitHub Actions | Run all checks on PR/push | `.github/workflows/ci.yml` |
| npm scripts | Orchestrate all tools | `package.json` |

## Tooling Data Flow

```
Developer writes code
        |
        v
npm run lint       -->  Biome checks style + correctness
npm run typecheck  -->  tsc verifies types (no emit)
npm run test       -->  Vitest runs tests + coverage
        |
        v
npm run check      -->  Runs all three in sequence (local)
        |
        v
git push / PR
        |
        v
GitHub Actions     -->  npm ci -> lint -> typecheck -> test
        |
        v
Pass/Fail status on PR
```

## Test Architecture

The server is a single file (~700 LOC), which means testing requires extracting testable units or testing through the public interfaces.

### Recommended Test Strategy

**Unit tests (high value, low effort):**
- MCP tool handler functions (if extractable)
- Command protocol message construction/parsing
- Parameter validation (Zod schemas)

**Integration-style tests (medium value, medium effort):**
- MCP tool registration (verify tools are exposed correctly)
- Command-response round trip (mock WebSocket)

**Not worth testing (low value, high effort):**
- HTTPS server TLS handshake
- Real WebSocket connection lifecycle
- PowerPoint add-in interaction (requires actual PowerPoint)

### Code Extraction Pattern

The current `server/index.ts` likely has tool handlers inline. For testability, consider extracting:

```typescript
// server/tools.ts - Extracted tool handlers (pure functions)
export function buildAddSlideCommand(params: AddSlideParams): WSCommand { ... }
export function buildAddShapeCommand(params: AddShapeParams): WSCommand { ... }
export function parseResponse(msg: WSMessage): ToolResult { ... }

// server/index.ts - Server wiring (imports and uses tool handlers)
```

This extraction is optional for initial release but recommended for test quality. Testing extracted pure functions is trivial; testing inline handlers in a monolithic server file requires more complex mocking.

## Patterns to Follow

### Pattern 1: Script-Based Tool Orchestration
**What:** All tools invoked through npm scripts, not custom build pipelines.
**Why:** Simple, portable, familiar to contributors. No Makefile, no task runner.
**Example:** `npm run check` runs lint + typecheck + test.

### Pattern 2: Fail-Fast CI Ordering
**What:** CI runs cheapest checks first: lint (seconds) -> typecheck (seconds) -> test (seconds-minutes).
**Why:** Fast feedback on obvious issues. No point running slow tests if code doesn't lint.

### Pattern 3: Config Files at Root
**What:** `biome.json` and `vitest.config.ts` at project root.
**Why:** Convention. Tools look for config at root by default. Contributors expect it there.

## Anti-Patterns to Avoid

### Anti-Pattern 1: Build Step for Tests
**What:** Adding a compile step before tests run.
**Why bad:** Node 24 runs TS natively. Vitest handles TS natively. Adding `tsc` build before test adds latency and a `dist/` directory to manage.
**Instead:** Run tests directly on `.ts` source files.

### Anti-Pattern 2: Separate Format/Lint Steps in CI
**What:** Running `biome format --check` and `biome lint` as separate CI steps.
**Why bad:** `biome check` does both in one pass. Two steps waste time.
**Instead:** Single `npm run lint` that runs `biome check .`.

### Anti-Pattern 3: Test Directory Mirroring Source
**What:** `tests/server/index.test.ts` mirroring `server/index.ts`.
**Why bad:** Single source file. Mirror structure adds unnecessary nesting.
**Instead:** Flat `tests/` directory with descriptive file names: `tools.test.ts`, `protocol.test.ts`.

## Sources

- Project structure from filesystem inspection
- [Vitest guide](https://vitest.dev/guide/) - test configuration patterns
- [Biome configuration](https://biomejs.dev/guides/configure-biome/) - config file placement
