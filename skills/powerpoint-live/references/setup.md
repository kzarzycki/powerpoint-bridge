# Setup

## Plugin install (recommended)

If you installed powerpoint-bridge as a Claude Code plugin, everything is automatic:
- MCP server starts via stdio when Claude Code launches
- Bridge server starts in the same process (add-in connection on port 8080)
- Add-in manifest is auto-sideloaded to PowerPoint on first run

Verify by calling `list_presentations`. If it returns results or "No presentations connected", setup is complete.

## Per-project setup (standalone)

For standalone users without the plugin, run these steps when asked to enable PowerPoint MCP in a project. Each step is idempotent — check before acting, skip what's already done.

### 1. Check if already configured

Look for `powerpoint-bridge` in the project's `.mcp.json`. If it exists, skip to step 3 (verify).

### 2. Add MCP config

Create or merge into the project's `.mcp.json`:

```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "command": "node",
      "args": ["<path-to-powerpoint-bridge>/server/index.ts", "--stdio", "--bridge"]
    }
  }
}
```

Replace `<path-to-powerpoint-bridge>` with the absolute path to the repo. If `.mcp.json` already exists with other servers, merge — do not overwrite.

### 3. Verify connectivity

Call `list_presentations`. If it returns results or "No presentations connected", setup is complete. Report status to user.
