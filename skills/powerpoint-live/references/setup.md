# Per-Project Setup

Run these steps when asked to enable PowerPoint MCP in a project. Each step is idempotent — check before acting, skip what's already done.

## 1. Check if already configured

Look for `powerpoint-bridge` in the project's `.mcp.json`. If it exists and points to `http://localhost:3001/mcp`, skip to step 4 (verify).

## 2. Add MCP config

Create or merge into the project's `.mcp.json`:

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

If `.mcp.json` already exists with other servers, merge — do not overwrite.

## 3. Check server is running

```bash
curl -s http://localhost:3001/mcp -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'
```

- If responds with JSON containing `serverInfo` — server is running, continue.
- If fails — tell user: "Start the bridge server: `cd ~/powerpoint-bridge && npm start`"

## 4. Verify connectivity

Call `list_presentations`. If it returns results or "No presentations connected", setup is complete. Report status to user.

## Not installed at all?

If the bridge repo doesn't exist at `~/powerpoint-bridge`, direct the user to the [README install instructions](../../../README.md#install).
