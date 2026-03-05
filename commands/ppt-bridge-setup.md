---
description: Check and configure the PowerPoint bridge (certs, manifest, server)
---

# PowerPoint Bridge Setup

You are running the setup check for the PowerPoint MCP bridge. Follow each step in order, report status, and stop at the first issue that needs user action.

## Step 1: Check TLS certificates

Check if certificate files exist:

```bash
ls "${CLAUDE_PLUGIN_ROOT}/certs/localhost.pem" "${CLAUDE_PLUGIN_ROOT}/certs/localhost-key.pem" 2>/dev/null
```

- If both files exist: report "Certs: OK" and continue.
- If missing: tell the user to generate them:
  ```
  cd <plugin-root> && npx mkcert -install && npx mkcert -key-file certs/localhost-key.pem -cert-file certs/localhost.pem localhost 127.0.0.1
  ```
  Note: `mkcert -install` requires an interactive terminal for the macOS password prompt. Stop here and wait for the user.

## Step 2: Check add-in manifest sideload

Check if the manifest is sideloaded:

```bash
ls ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef/ 2>/dev/null | grep -q .
```

- If the wef directory has files: report "Manifest: OK (sideloaded)" and continue.
- If empty or missing: sideload it:
  ```bash
  mkdir -p ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef
  cp "${CLAUDE_PLUGIN_ROOT}/addin/manifest.xml" ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef/
  ```
  Report "Manifest: sideloaded" and continue.

## Step 3: Check bridge server

Test if the bridge server is running:

```bash
curl -sf http://localhost:3001/health
```

- If responds with JSON containing `"status":"ok"`: report "Server: running (N presentations connected)" using the `connections` value from the response, and continue.
- If fails: tell the user to start it:
  ```
  cd <plugin-root> && npm start
  ```
  Stop here and wait for the user.

## Step 4: Report summary

Print a summary table:

```
PowerPoint Bridge Status
========================
Certs:    OK
Manifest: OK (sideloaded)
Server:   running

Ready to use! Open a presentation in PowerPoint and the add-in will connect automatically.
```

If any step failed, only show status up to the failure point with a clear next action.
