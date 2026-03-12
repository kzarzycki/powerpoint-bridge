## GitHub Actions Workflow Model

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
