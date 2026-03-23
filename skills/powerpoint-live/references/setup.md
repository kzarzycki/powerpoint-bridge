# Setup

## Plugin install (recommended)

Install as a Claude Code plugin — this is the simplest way to get started:

```bash
claude plugin add kzarzycki/powerpoint-bridge
```

This gives you:
- MCP server starts via stdio when Claude Code launches
- Bridge server starts in the same process (add-in connection on port 8080)
- Skills auto-discovered by the plugin system

### Sideload the PowerPoint add-in

The plugin handles the MCP server, but PowerPoint still needs the add-in manifest sideloaded once:

```bash
# From the plugin's repo directory:
node scripts/sideload.mjs
```

Then restart PowerPoint to load the add-in.

### Verify

Call `list_presentations`. If it returns results or "No presentations connected", setup is complete.

## Standalone setup (for developers)

If you're developing on the bridge itself or prefer not to use the plugin:

### 1. Clone and install

```bash
git clone https://github.com/kzarzycki/powerpoint-bridge.git
cd powerpoint-bridge
npm install
```

### 2. Sideload the add-in

```bash
npm run sideload
```

Restart PowerPoint after sideloading.

### 3. Start the bridge server

Start the server once in HTTP mode — it stays running across Claude Code sessions:

```bash
nohup node --experimental-strip-types ./server/index.ts --http --bridge > /tmp/powerpoint-bridge.log 2>&1 &
```

To restart after code changes:

```bash
pkill -f "server/index.ts"
nohup node --experimental-strip-types ./server/index.ts --http --bridge > /tmp/powerpoint-bridge.log 2>&1 &
```

HTTP mode is preferred over STDIO for development because:
- Multiple Claude Code sessions can connect simultaneously
- Claude can autonomously restart the server via Bash (closes the dev feedback loop)
- No build step needed — runs directly from TypeScript source

### 4. Add MCP config to your project

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

<details>
<summary>Alternative: STDIO mode (single session)</summary>

```json
{
  "mcpServers": {
    "powerpoint-bridge": {
      "command": "node",
      "args": ["<path-to-powerpoint-bridge>/dist/index.cjs", "--stdio", "--bridge"]
    }
  }
}
```

STDIO ties the server lifecycle to Claude Code — you can't restart it independently.
</details>

### 5. Verify

Call `list_presentations`. If it returns results or "No presentations connected", setup is complete.
