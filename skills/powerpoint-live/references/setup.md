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
curl -sf http://localhost:3001/health
```

- If responds with JSON containing `"status":"ok"` — server is running, continue to step 4.
- If fails — tell user: "The bridge server isn't running. Start it with `npm start` from the powerpoint-bridge repo directory." Optionally suggest setting up a launchd plist for persistent auto-start if the user wants the server to run automatically.

## 4. Verify connectivity

Call `list_presentations`. If it returns results or "No presentations connected", setup is complete. Report status to user.
