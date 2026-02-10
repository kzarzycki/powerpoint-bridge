# Milestone v2: Open Source Release

**Status:** SHIPPED 2026-02-10
**Phases:** 6-9
**Total Plans:** 4 (executed directly, not via GSD phase workflow)

## Overview

Make the project a proper open-source repository that anyone with macOS + PowerPoint can clone, set up, and use — with clean code, tests, docs, and CI. First public release as v0.1.0.

## Phases

### Phase 6: Repo Hygiene

**Goal:** Legal, metadata, and path cleanup so the repo is publishable.

- [x] Create MIT LICENSE
- [x] Expand .gitignore (macOS, IDE, coverage, build artifacts)
- [x] Add package.json metadata (description, license, repository, keywords, engines, private)
- [x] Set version to 0.1.0
- [x] Fix CLAUDE.md architecture diagram (MCP stdio → MCP HTTP)
- [x] Replace hardcoded /Users/zarz/ paths in .planning/ files

### Phase 7: Code Quality

**Goal:** Split server for testability, add lint + format.

- [x] Refactor server/index.ts (558 LOC) into bridge.ts + tools.ts + index.ts
- [x] ConnectionPool class replaces module-level state
- [x] registerTools() receives dependencies as args
- [x] Install and configure Biome (lint + format)
- [x] Auto-fix all existing code
- [x] Add lint, lint:fix, format, check scripts

### Phase 8: Testing

**Goal:** Meaningful test coverage of the testable core.

- [x] Install Vitest + @vitest/coverage-v8
- [x] bridge.test.ts: 15 tests (ConnectionPool, resolveTarget, sendCommand, timeout)
- [x] tools.test.ts: 7 tests via InMemoryTransport (tool listing, list_presentations, execute_officejs)
- [x] Coverage: bridge.ts 96%, tools.ts 59%

### Phase 9: CI & Documentation

**Goal:** Automated quality gates and human-readable docs.

- [x] GitHub Actions CI (ubuntu-latest, Node 24, lint → typecheck → test)
- [x] README.md (architecture, prerequisites, quick start, MCP configs, tools, limitations, security, troubleshooting)
- [x] CONTRIBUTING.md (dev setup, structure, scripts, PR process)

## Milestone Summary

**Key Decisions:**
- Biome over ESLint+Prettier (single tool, Rust-based, simpler)
- 3-file server split over monolith (testability without over-engineering)
- Vitest over Jest/node:test (native ESM+TS, MCP SDK uses it)
- Keep .planning/ and git history (educational value)

**Issues Resolved:**
- Hardcoded user paths in 12 files
- Untestable monolithic server (558 LOC with module-level side effects)
- No quality gates (lint, typecheck, test were all manual)

**Technical Debt Incurred:**
- tools.ts coverage at 59% (get_slide handler untested)
- Ports hardcoded (8443, 3001) — todo captured for future configurable ports with add-in UI

---

_For current project status, see .planning/ROADMAP.md_
