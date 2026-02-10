# Contributing to PowerPoint Bridge

## Prerequisites

- macOS with PowerPoint for Mac installed
- Node.js >= 24
- mkcert (`brew install mkcert`)

## Development Setup

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git
cd powerpoint-bridge
npm install

# Generate TLS certs (one-time)
mkcert -install
npm run setup-certs

# Sideload the add-in
npm run sideload
# Restart PowerPoint after sideloading
```

## Project Structure

```
server/
  index.ts      # Entrypoint: HTTPS + WSS + MCP servers, wiring
  bridge.ts     # ConnectionPool class — manages add-in WebSocket connections
  tools.ts      # MCP tool definitions (list_presentations, get_presentation, etc.)
  bridge.test.ts
  tools.test.ts
addin/
  index.html    # Add-in taskpane UI
  app.js        # WebSocket client + Office.js command execution
  manifest.xml  # Office Add-in manifest for sideloading
certs/          # Generated TLS certs (gitignored)
```

## Scripts

| Script | Purpose |
|--------|---------|
| `npm start` | Start the bridge server |
| `npm test` | Run tests |
| `npm run test:watch` | Run tests in watch mode |
| `npm run test:coverage` | Run tests with coverage report |
| `npm run lint` | Check code with Biome |
| `npm run lint:fix` | Auto-fix lint and format issues |
| `npm run typecheck` | TypeScript type checking |
| `npm run check` | Run lint + typecheck + tests (CI equivalent) |

## Code Style

[Biome](https://biomejs.dev/) enforces formatting and lint rules. Run `npm run lint:fix` before committing. Key settings:

- 2 spaces, no semicolons, single quotes
- 120 character line width

## Running Tests

Tests use [Vitest](https://vitest.dev/) and don't require macOS, PowerPoint, or TLS certificates:

```bash
npm test              # Single run
npm run test:watch    # Watch mode
npm run test:coverage # With coverage
```

## Pull Request Process

1. Create a feature branch from `main`
2. Make your changes
3. Run `npm run check` to verify lint, types, and tests pass
4. Open a PR against `main`

## What's In Scope

- Office.js capabilities within Requirement Sets 1.1-1.9
- MCP tool improvements
- macOS support improvements
- Documentation

## What's Out of Scope

- Features requiring Office.js APIs not available on Mac (e.g., images, charts)
- Windows/Linux platform support (PRs welcome, but we can't test)
- Non-PowerPoint Office apps
