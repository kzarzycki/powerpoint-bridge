# Technology Stack: Open-Source Release Tooling

**Project:** PowerPoint Office.js Bridge
**Researched:** 2026-02-10
**Scope:** Linting, testing, CI/CD, and code quality tooling for OSS release

## Recommended Stack

### Linting & Formatting: Biome

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| `@biomejs/biome` | ^2.3.14 | Linting + formatting (unified) | Single tool replaces ESLint + Prettier. 10-25x faster. One config file. Zero JS dependencies (ships as native binary). Proportional to a 700 LOC project -- no config sprawl. |

**Confidence:** HIGH (verified via [biomejs.dev](https://biomejs.dev/), [npm registry](https://www.npmjs.com/package/@biomejs/biome), [GitHub releases](https://github.com/biomejs/biome/releases))

**Why not ESLint + Prettier:**
- ESLint v9.39.2 + typescript-eslint v8.55.0 + Prettier requires 4+ packages and 3+ config files
- ESLint v10.0.0-rc.2 is in release candidate phase -- adopting now means migration soon
- For a 700 LOC single-server project, the ESLint plugin ecosystem is overkill
- Biome covers 434+ rules including TypeScript-specific ones, which is more than enough

**Why not ESLint alone (without Prettier):**
- ESLint deprecated all formatting rules in v8.53.0; formatting requires Prettier or stylistic plugin
- Adding `@stylistic/eslint-plugin` is another dependency; Biome handles both natively

**Recommended `biome.json`:**

```json
{
  "$schema": "https://biomejs.dev/schemas/2.3.14/schema.json",
  "vcs": {
    "enabled": true,
    "clientKind": "git",
    "useIgnoreFile": true
  },
  "files": {
    "includes": ["server/**/*.ts", "addin/**/*.js"]
  },
  "formatter": {
    "indentStyle": "space",
    "indentWidth": 2,
    "lineWidth": 100
  },
  "linter": {
    "enabled": true,
    "rules": {
      "recommended": true
    }
  },
  "javascript": {
    "formatter": {
      "quoteStyle": "single",
      "semicolons": "asNeeded"
    }
  }
}
```

**Scripts to add:**

```json
{
  "lint": "biome check .",
  "lint:fix": "biome check --write .",
  "format": "biome format --write ."
}
```

---

### Testing: Vitest

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| `vitest` | ^4.0.18 | Test runner + assertions + mocking | Native TypeScript + ESM support. Zero-config for this project shape. Jest-compatible API (familiar to contributors). Built-in watch mode and coverage. |
| `@vitest/coverage-v8` | ^4.0.18 | Code coverage via V8 engine | Built-in V8 coverage provider. No separate istanbul config needed. Accurate AST-based remapping in v4. |

**Confidence:** HIGH (verified via [vitest.dev](https://vitest.dev/guide/), [npm registry](https://www.npmjs.com/package/vitest), [GitHub releases](https://github.com/vitest-dev/vitest/releases))

**Why not Node.js built-in test runner (`node:test`):**
- `node:test` works for basic cases but has limited mocking helpers, no built-in watch mode, and a less familiar API
- Testing WebSocket connections and MCP protocol handlers benefits from Vitest's mocking and module interception
- Contributors expect Vitest or Jest -- `node:test` has lower recognition in OSS
- Coverage with `node:test` requires separate tooling; Vitest has it built in

**Why not Jest:**
- Jest 30 (June 2025) improved ESM support but it's still experimental
- Jest requires `ts-jest` or Babel for TypeScript -- Vitest handles it natively
- Vitest runs 10-20x faster in watch mode
- For a new project with no existing Jest tests, Vitest is the modern default

**Recommended `vitest.config.ts`:**

```typescript
import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      include: ['server/**/*.ts'],
      exclude: ['server/**/*.d.ts'],
      reporter: ['text', 'lcov'],
      thresholds: {
        statements: 60,
        branches: 60,
        functions: 60,
        lines: 60
      }
    }
  }
})
```

**Scripts to add:**

```json
{
  "test": "vitest run",
  "test:watch": "vitest",
  "test:coverage": "vitest run --coverage"
}
```

---

### CI/CD: GitHub Actions

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| `actions/checkout` | v6 | Clone repository | Standard, current version |
| `actions/setup-node` | v6 | Install Node.js 24 | Supports Node 24, built-in npm caching |
| GitHub Actions workflow | N/A | CI pipeline | Free for public repos. Industry standard for OSS. |

**Confidence:** HIGH (verified via [actions/setup-node](https://github.com/actions/setup-node), [GitHub docs](https://docs.github.com/en/actions/automating-builds-and-tests/building-and-testing-nodejs))

**Recommended `.github/workflows/ci.yml`:**

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

permissions:
  contents: read

jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
      - run: npm ci
      - run: npm run lint
      - run: npm run typecheck
      - run: npm test
```

**Notes:**
- Single job is appropriate for a 700 LOC project (no need for matrix or parallel jobs)
- Node 24 only -- no matrix testing needed since the project explicitly requires Node 24 for native TS
- `npm ci` for reproducible installs from lockfile
- Order: lint -> typecheck -> test (fail fast on cheapest checks first)

---

### Type Checking: TypeScript (already present)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| `typescript` | ^5.9.0 | Static type checking | Already configured. `tsc --noEmit` via `npm run typecheck`. No changes needed. |

**Confidence:** HIGH (already in project, verified working)

The existing `tsconfig.json` is well-configured for Node 24 native TS execution:
- `erasableSyntaxOnly: true` enforces Node-compatible TS syntax
- `noEmit: true` since Node runs TS directly
- `strict: true` for type safety

---

## What NOT to Add

These tools are commonly recommended but would be over-engineering for this project.

| Tool | Why Skip |
|------|----------|
| **Husky / Lefthook** (pre-commit hooks) | 700 LOC, single developer. CI catches everything. Add only if contribution volume grows. |
| **lint-staged** | Only useful with pre-commit hooks (see above). |
| **Changesets / semantic-release** | Premature for initial release. Manual versioning with `npm version` is sufficient. Revisit after v1.0. |
| **Commitlint / Conventional Commits** | Adds friction for contributors without proportional benefit at this scale. |
| **Prettier** | Biome handles formatting. Adding Prettier alongside Biome creates conflicts. |
| **c8 / istanbul** (standalone) | Vitest has built-in coverage. No separate tool needed. |
| **tsx / ts-node** | Node 24 runs TypeScript natively. These are unnecessary. |
| **Turborepo / Nx** | Single-package project. Monorepo tooling is irrelevant. |

---

## Alternatives Considered

| Category | Recommended | Alternative | Why Not Alternative |
|----------|-------------|-------------|---------------------|
| Lint + Format | Biome 2.3 | ESLint 9 + Prettier | 4+ packages, 3+ config files, ESLint 10 migration imminent |
| Testing | Vitest 4.0 | Jest 30 | ESM still experimental in Jest, needs ts-jest, slower |
| Testing | Vitest 4.0 | node:test | Limited mocking, no watch, unfamiliar to contributors |
| CI | GitHub Actions | CircleCI / Travis | GitHub Actions is free for public repos, best integration |
| Formatting | Biome | dprint | Less community adoption, separate from linting |

---

## Installation

```bash
# Dev dependencies for quality tooling
npm install -D @biomejs/biome vitest @vitest/coverage-v8

# That's it. Three packages total for lint + format + test + coverage.
```

**Total new devDependencies: 3**

Compare with ESLint + Prettier + Jest equivalent:
- eslint, @eslint/js, typescript-eslint, prettier, eslint-config-prettier, jest, ts-jest, @types/jest, @jest/globals = 9+ packages

---

## Package.json Scripts (Complete)

After adding tooling, the scripts section should be:

```json
{
  "scripts": {
    "start": "node server/index.ts",
    "typecheck": "tsc --noEmit",
    "lint": "biome check .",
    "lint:fix": "biome check --write .",
    "format": "biome format --write .",
    "test": "vitest run",
    "test:watch": "vitest",
    "test:coverage": "vitest run --coverage",
    "check": "npm run lint && npm run typecheck && npm test",
    "setup-certs": "mkdir -p certs && mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1",
    "sideload": "mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef && cp addin/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/"
  }
}
```

The `check` script runs all quality gates in sequence -- useful for local pre-push verification without needing git hooks.

---

## Sources

- [Biome official site](https://biomejs.dev/) - v2.3.14, configuration reference
- [Biome GitHub releases](https://github.com/biomejs/biome/releases) - version verification
- [Vitest official docs](https://vitest.dev/guide/) - v4.0.18, configuration
- [Vitest coverage guide](https://vitest.dev/guide/coverage) - v8 provider setup
- [actions/setup-node](https://github.com/actions/setup-node) - v6, Node 24 support
- [GitHub Actions Node.js docs](https://docs.github.com/en/actions/automating-builds-and-tests/building-and-testing-nodejs) - CI best practices
- [typescript-eslint](https://typescript-eslint.io/) - v8.55.0 (considered, not recommended)
- [ESLint releases](https://github.com/eslint/eslint/releases) - v9.39.2, v10.0.0-rc.2
- [Node.js TypeScript docs](https://nodejs.org/en/learn/typescript/run-natively) - native type stripping
