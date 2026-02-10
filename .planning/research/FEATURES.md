# Feature Landscape: Open-Source Release Quality

**Domain:** OSS release readiness for a Node.js/TypeScript MCP server
**Researched:** 2026-02-10

## Table Stakes

Features users/contributors expect. Missing = project feels amateurish.

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| Linting with auto-fix | Contributors need consistent code style feedback | Low | Biome `check --write` |
| Formatting | PRs shouldn't have style debates | Low | Biome handles this alongside linting |
| Type checking | TypeScript project must pass `tsc --noEmit` | Low | Already exists (`npm run typecheck`) |
| Basic test suite | Demonstrates project works, enables contributions | Medium | Vitest for MCP tools, command protocol |
| CI pipeline | PRs must be gated on quality checks | Low | GitHub Actions, single workflow |
| README with setup instructions | First thing visitors see | Low | Already partially exists in CLAUDE.md |
| LICENSE file | Legal requirement for OSS use | Low | MIT is standard for this type of project |
| .gitignore covering common patterns | Prevents committing build artifacts, secrets | Low | Already exists, may need expansion |
| package.json with correct metadata | npm/GitHub display, discoverability | Low | Add description, repository, keywords, license fields |

## Differentiators

Features that signal quality. Not expected, but valued by contributors.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| Code coverage reporting | Shows which code paths are tested | Low | Built into Vitest with v8 provider |
| CONTRIBUTING.md | Reduces friction for first-time contributors | Low | Brief: setup, code style, PR process |
| Inline JSDoc on public API | Helps users understand MCP tools without reading source | Low-Medium | Focus on exported functions and types |
| Error handling documentation | MCP tools should document failure modes | Low | In README or code comments |
| Example usage / demo | Shows the product working | Medium | Could be a GIF in README |

## Anti-Features

Features to explicitly NOT build for initial release. Common mistakes in OSS.

| Anti-Feature | Why Avoid | What to Do Instead |
|--------------|-----------|-------------------|
| 100% test coverage | Diminishing returns on I/O-heavy code. Tests become brittle mocks. | Target 60% on critical paths (MCP tools, command protocol). |
| Pre-commit hooks (Husky/Lefthook) | Friction for single-developer project. CI catches issues. | Add `npm run check` script for manual pre-push verification. |
| Automated changelog generation | No release history exists yet. Premature automation. | Manual CHANGELOG.md entries when releasing. |
| Monorepo tooling | Single package with ~700 LOC. | Keep flat structure. |
| API documentation site | No external API consumers beyond MCP. | Inline JSDoc + README. |
| Docker containerization | macOS-only project (requires PowerPoint + mkcert). | Document native setup only. |
| npm publishing | This is a CLI tool, not a library. | Users clone the repo and run `npm start`. |
| Semantic versioning automation | No release cadence established. | `npm version` manually. |

## Feature Dependencies

```
LICENSE (independent, do first)
    |
Biome config --> Lint fix existing code --> CI workflow
    |                                          |
Vitest config --> Write tests ---------> CI includes tests
    |                                          |
README rewrite <-- all tooling configured <----+
    |
CONTRIBUTING.md (references README setup)
```

## MVP Recommendation

For initial OSS release, prioritize:

1. LICENSE (MIT) -- legal foundation, 5 minutes
2. Biome setup + lint fix -- establishes code quality baseline
3. Vitest + core tests -- proves the project works
4. CI workflow -- automates quality gates
5. README rewrite -- converts CLAUDE.md content to contributor-friendly format

Defer to post-release:
- CONTRIBUTING.md: Add after first external contribution request
- Coverage badges: Add after coverage stabilizes
- Demo GIF/video: Add after gathering user interest
- Changelog automation: Add after v1.0 with regular releases

## Sources

- Community patterns from [ESLint](https://github.com/eslint/eslint), [Vitest](https://github.com/vitest-dev/vitest), and similar OSS projects
- [GitHub docs on open source](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features)
