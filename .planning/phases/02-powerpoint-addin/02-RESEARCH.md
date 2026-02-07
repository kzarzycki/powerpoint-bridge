# Phase 2: PowerPoint Add-in - Research

**Researched:** 2026-02-07
**Domain:** Office.js add-in development, XML manifest, WKWebView WebSocket, macOS sideloading
**Confidence:** HIGH

## Summary

This phase creates the Office.js add-in that loads inside PowerPoint's taskpane and connects via WebSocket to the bridge server from Phase 1. The add-in is pure HTML/CSS/JS (no build step), served as static files from the existing HTTPS server on port 8443, and loaded into PowerPoint via a sideloaded XML manifest.

The critical technical points are: (1) the XML manifest format is required (unified JSON manifest cannot be sideloaded on Mac), (2) Office.js must be loaded from Microsoft's CDN in the `<head>`, (3) WSS is mandatory in WKWebView (mkcert certificates are system-trusted via Keychain so WSS works), and (4) the WebSocket client needs exponential backoff reconnection since the add-in may load before the bridge server is running.

**Primary recommendation:** Build a minimal HTML taskpane with Office.js CDN initialization and a vanilla WebSocket client with reconnection logic. Use the XML manifest targeting the existing `https://localhost:8443` server. Icon images should be served from the same HTTPS server.

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| Office.js (CDN) | 1.1 (latest) | Office add-in API bridge | Only way to interact with PowerPoint from a taskpane; loaded from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` |
| Native WebSocket API | (browser built-in) | WSS client connection | Built into WKWebView/Safari; no library needed for basic WebSocket client |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| @types/office-js | latest | TypeScript type definitions | Only for type-checking during development; the add-in itself is plain JS |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| CDN Office.js | npm @microsoft/office-js | CDN is required for production; npm package exists but CDN is the official approach |
| Native WebSocket | reconnecting-websocket library | Adds a dependency; reconnection is simple enough to hand-roll for this use case |
| XML manifest | Unified JSON manifest | JSON manifest is in preview for PowerPoint and CANNOT be sideloaded on Mac |

**Installation:**
```bash
# No new npm packages needed for the add-in itself
# For TypeScript type-checking only:
npm install --save-dev @types/office-js
```

## Architecture Patterns

### Recommended Project Structure
```
addin/
  index.html          # Main taskpane HTML (Office.js init + WS client + UI)
  style.css           # Taskpane styles (optional, can be inline)
  app.js              # WebSocket client logic and UI updates
  assets/
    icon-16.png       # Ribbon icon 16x16
    icon-32.png       # Ribbon icon 32x32
    icon-80.png       # Ribbon icon 80x80
  manifest.xml        # XML manifest for sideloading (copied to wef/)
```

### Pattern 1: Office.js Initialization with onReady
**What:** Use `Office.onReady()` to detect PowerPoint context before initializing the WebSocket connection.
**When to use:** Always -- this is the entry point for any Office add-in code.
**Example:**
```javascript
// Source: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    // We are running inside PowerPoint
    updateStatus('Office.js ready');
    initWebSocket();
  }
  if (info.platform === Office.PlatformType.Mac) {
    // macOS-specific adjustments if needed
  }
});
```

### Pattern 2: WebSocket Client with Exponential Backoff Reconnection
**What:** A WebSocket client that auto-reconnects with increasing delays when disconnected.
**When to use:** Always -- the add-in may load before the server, or the server may restart.
**Example:**
```javascript
// Reconnection with exponential backoff + jitter
const WS_URL = 'wss://localhost:8443';
let ws = null;
let reconnectAttempt = 0;
const BASE_DELAY = 500;    // ms
const MAX_DELAY = 30000;   // ms

function connect() {
  ws = new WebSocket(WS_URL);

  ws.onopen = function() {
    reconnectAttempt = 0;
    updateStatus('Connected');
  };

  ws.onclose = function() {
    updateStatus('Disconnected');
    scheduleReconnect();
  };

  ws.onerror = function() {
    // onclose will fire after onerror, so reconnect happens there
  };

  ws.onmessage = function(event) {
    // Handle incoming commands from bridge server
    var message = JSON.parse(event.data);
    handleCommand(message);
  };
}

function scheduleReconnect() {
  var delay = Math.min(BASE_DELAY * Math.pow(2, reconnectAttempt), MAX_DELAY);
  var jitter = Math.floor(Math.random() * 1000);
  reconnectAttempt++;
  setTimeout(connect, delay + jitter);
}
```

### Pattern 3: XML Manifest with VersionOverrides for Ribbon Button
**What:** The manifest defines a custom ribbon button in the Home tab that opens the taskpane.
**When to use:** Always -- this is how the add-in integrates into PowerPoint's UI.

### Anti-Patterns to Avoid
- **Loading Office.js from npm bundle instead of CDN:** The CDN URL is the official and required approach. Loading from a local bundle can cause version mismatches with the host application.
- **Using Office.initialize instead of Office.onReady:** `Office.initialize` is legacy, can only have one handler, must be assigned before the event fires. `Office.onReady()` is the modern approach -- returns a Promise, can be called multiple times.
- **Opening WebSocket before Office.onReady:** The add-in HTML may load in a normal browser context (for testing). Check that Office.js is ready before assuming the PowerPoint context exists. However, the WebSocket connection itself can start independently.
- **Putting manifest.xml inside addin/ directory for sideloading:** The manifest must be copied to the wef directory. The file in addin/ is the source; sideloading requires it in `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Office.js loading | Custom loader or npm bundle | CDN script tag in `<head>` | CDN auto-updates, stays compatible with host app version |
| GUID generation for manifest | Manual UUID | `uuidgen` CLI or online generator | Manifest ID must be a valid GUID; generate once, never changes |
| Icon creation | Complex graphics | Simple solid-color PNG with letter/symbol | Office only needs 16, 32, 80px PNGs; elaborate icons waste time for a dev tool |
| WebSocket protocol | Custom binary protocol | JSON text messages | Simple, debuggable, matches the command protocol already designed in CLAUDE.md |

**Key insight:** The add-in is intentionally minimal -- its only job is to connect via WebSocket and relay commands. Keep the HTML/CSS/JS as simple as possible. No framework, no bundler, no build step.

## Common Pitfalls

### Pitfall 1: Manifest Element Ordering
**What goes wrong:** The XML manifest schema enforces strict element ordering. Elements out of order cause cryptic validation errors, and the add-in silently fails to load.
**Why it happens:** XML schemas use `xs:sequence` which mandates order. The error messages don't clearly state "wrong order."
**How to avoid:** Follow the exact ordering from Microsoft's documentation. The order is: Id, AlternateId, Version, ProviderName, DefaultLocale, DisplayName, Description, IconUrl, HighResolutionIconUrl, SupportUrl, AppDomains, Hosts, Requirements, DefaultSettings, Permissions, VersionOverrides.
**Warning signs:** Add-in doesn't appear in PowerPoint's ribbon after sideloading; no error message visible.

### Pitfall 2: wef Directory Doesn't Exist
**What goes wrong:** The `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` directory may not exist if no add-in has ever been sideloaded.
**Why it happens:** PowerPoint creates this directory lazily, not on install.
**How to avoid:** Always `mkdir -p` the directory before copying the manifest.
**Warning signs:** `cp` command fails silently or throws "No such file or directory."

### Pitfall 3: PowerPoint Must Be Restarted After Sideloading
**What goes wrong:** Copying the manifest to wef/ while PowerPoint is running does not immediately make the add-in available.
**Why it happens:** PowerPoint reads the wef/ directory on startup.
**How to avoid:** Close and reopen PowerPoint after placing the manifest. The add-in then appears under Home > Add-ins.
**Warning signs:** Manifest is in the right place but add-in doesn't appear in ribbon.

### Pitfall 4: Office.js CDN Script Must Be in Head
**What goes wrong:** Placing the Office.js `<script>` tag in `<body>` or loading it asynchronously can cause `Office.onReady()` to never resolve or to miss the initialization window.
**Why it happens:** Office.js must initialize before any body elements load.
**How to avoid:** Always place `<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js">` in the `<head>` section, before any other scripts.
**Warning signs:** `Office.onReady()` callback never fires; add-in shows blank taskpane.

### Pitfall 5: WSS Certificate Not Trusted by WKWebView
**What goes wrong:** WebSocket connection fails with a security error because the self-signed certificate is not trusted by the system.
**Why it happens:** WKWebView uses the macOS system trust store. If `mkcert -install` was not run (which adds the local CA to Keychain), the cert is untrusted.
**How to avoid:** Ensure `mkcert -install` was run during Phase 1 setup. This is a prerequisite. The mkcert local CA in the system Keychain is what makes WKWebView trust the localhost certificate.
**Warning signs:** WebSocket `onerror` fires immediately on connection attempt; console shows SSL/TLS errors.

### Pitfall 6: Personality Menu Obscures Top-Right UI
**What goes wrong:** On macOS, the personality menu (user identity icon) occupies 26x26 pixels in the top-right corner of the taskpane, offset 8px from right and 6px from top.
**Why it happens:** Office injects this control into the taskpane viewport.
**How to avoid:** Don't place important UI elements in the top-right 34x32 pixel area of the taskpane.
**Warning signs:** UI elements are hidden behind the circular user icon.

## Code Examples

### Complete Taskpane HTML Structure
```html
<!-- Source: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn -->
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PowerPoint Bridge</title>
  <!-- Office.js MUST be in <head>, before any other scripts -->
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <h1>PowerPoint Bridge</h1>
  <div id="status">Initializing...</div>
  <script src="app.js"></script>
</body>
</html>
```

### Office.js Initialization
```javascript
// Source: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in
Office.onReady(function(info) {
  // info.host: Office.HostType.PowerPoint
  // info.platform: Office.PlatformType.Mac
  console.log('Office.js ready:', info.host, info.platform);
  document.getElementById('status').textContent = 'Connecting...';
  initWebSocket();
});
```

### Runtime Requirement Set Check
```javascript
// Source: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
  // Shape manipulation APIs available
}
```

### Complete XML Manifest for PowerPoint Taskpane Add-in
```xml
<!-- Source: https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/hello-world/powerpoint-hello-world/manifest.xml -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <!-- GUID: generate once with uuidgen, never changes -->
  <Id>XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>PowerPoint Bridge</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="PowerPoint Bridge"/>
  <Description DefaultValue="Bridge for Claude Code to manipulate live PowerPoint presentations"/>
  <IconUrl DefaultValue="https://localhost:8443/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:8443/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://localhost:8443"/>

  <AppDomains>
    <AppDomain>https://localhost:8443</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8443/index.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="BridgeGroup">
                <Label resid="GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon16"/>
                  <bt:Image size="32" resid="Icon32"/>
                  <bt:Image size="80" resid="Icon80"/>
                </Icon>
                <Control xsi:type="Button" id="BridgeButton">
                  <Label resid="ButtonLabel"/>
                  <Supertip>
                    <Title resid="ButtonLabel"/>
                    <Description resid="ButtonTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon16"/>
                    <bt:Image size="32" resid="Icon32"/>
                    <bt:Image size="80" resid="Icon80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>BridgeTaskpane</TaskpaneId>
                    <SourceLocation resid="TaskpaneUrl"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon16" DefaultValue="https://localhost:8443/assets/icon-16.png"/>
        <bt:Image id="Icon32" DefaultValue="https://localhost:8443/assets/icon-32.png"/>
        <bt:Image id="Icon80" DefaultValue="https://localhost:8443/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="TaskpaneUrl" DefaultValue="https://localhost:8443/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Bridge"/>
        <bt:String id="ButtonLabel" DefaultValue="PowerPoint Bridge"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="ButtonTooltip" DefaultValue="Open the PowerPoint Bridge taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

### Sideloading Script
```bash
# Create wef directory if it doesn't exist
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef

# Copy manifest
cp addin/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/

# Restart PowerPoint manually after this
echo "Manifest installed. Restart PowerPoint, then find 'PowerPoint Bridge' under Home > Add-ins"
```

## Taskpane Dimensions and Layout

| Property | Value | Source |
|----------|-------|--------|
| Default size (PowerPoint) | 348 x 391 px | Microsoft Learn taskpane docs |
| User resizable | Width only, can increase but not decrease below default | Microsoft Learn |
| Personality menu (Mac) | 26x26 px, offset 8px from right, 6px from top | Microsoft Learn |
| Safe zone (top-right) | Avoid 34x32 px in top-right corner | Calculated from personality menu |
| Recommended design | Narrow-first, ~300-350px width | Best practice |

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `Office.initialize` callback | `Office.onReady()` Promise | Office.js 1.1 update | onReady is preferred; returns Promise, callable multiple times |
| XML manifest only | XML manifest + Unified JSON manifest | 2023-2024 | JSON is preview-only for PowerPoint; XML required for Mac sideloading |
| Internet Explorer WebView (Windows) | WKWebView (Mac) / Edge WebView2 (Windows) | macOS always used WebKit | Mac add-ins run in Safari/WebKit engine with modern ES6+ support |

**Deprecated/outdated:**
- `Office.initialize`: Still works but `Office.onReady()` is recommended
- Unified JSON manifest for PowerPoint on Mac: Currently preview-only and cannot be sideloaded on Mac

## Open Questions

1. **Icon image generation**
   - What we know: Need 16x16, 32x32, and 80x80 PNG files served from HTTPS. Simple solid-color icons with a letter are fine for a dev tool.
   - What's unclear: Whether dynamically generated SVG or data URI icons could work instead of separate PNG files (manifest requires HTTPS URLs to image resources).
   - Recommendation: Create minimal PNG icons (solid color background with "PB" text). Can be created programmatically with a canvas script or manually. Keep it simple.

2. **WKWebView WSS trust with mkcert on macOS 15+**
   - What we know: mkcert adds its local CA to the macOS system Keychain. Safari/WKWebView uses the system trust store. The Phase 1 RESEARCH.md confirms "Full WebSocket API support confirmed" for WKWebView.
   - What's unclear: Some Apple Developer Forum posts suggest macOS 15 tightened WSS restrictions. The project's existing RESEARCH.md and the mkcert-for-Office-addins blog post both support that mkcert should work.
   - Recommendation: Proceed with mkcert WSS approach. If it fails at testing time, the error will be immediately visible ("Disconnected" status). Fallback: HTTP polling (much less desirable).

3. **FunctionFile element in VersionOverrides**
   - What we know: The `<FunctionFile>` element is shown in Microsoft's documentation examples within `<DesktopFormFactor>`. It points to an HTML file that contains JavaScript for `ExecuteFunction`-type commands.
   - What's unclear: Whether `<FunctionFile>` is required when we only use `ShowTaskpane` actions (no `ExecuteFunction` actions). The hello-world sample omits it; the template includes it.
   - Recommendation: Omit `<FunctionFile>` since we only have a `ShowTaskpane` button. The hello-world sample proves this works.

## Sources

### Primary (HIGH confidence)
- Microsoft Learn: [Referencing Office.js from CDN](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn) - CDN URL, head placement requirement
- Microsoft Learn: [Initialize Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in) - Office.onReady vs Office.initialize, HostType/PlatformType values
- Microsoft Learn: [Sideload on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac) - wef directory path, sideloading steps
- Microsoft Learn: [XML Manifest Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview) - Full manifest schema, element ordering, PowerPoint host specification
- Microsoft Learn: [Create add-in commands](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands) - VersionOverrides, DesktopFormFactor, ExtensionPoint, ShowTaskpane action
- Microsoft Learn: [Task pane design](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) - Taskpane dimensions, personality menu, layout guidance
- Microsoft Learn: [Manifest element ordering](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/manifest-element-ordering) - Strict element order requirements
- Microsoft Learn: [Specify requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) - Requirements element, PowerPointApi sets, isSetSupported()
- Microsoft Learn: [Icon guidelines](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-icons) - Required sizes, PNG format, color guidelines
- Microsoft Learn: [Clear cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache) - Cache clearing steps for macOS
- GitHub: [PowerPoint hello-world manifest.xml](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/hello-world/powerpoint-hello-world/manifest.xml) - Complete working manifest example
- GitHub: [office-js issue #1018](https://github.com/OfficeDev/office-js/issues/1018) - SecurityError with ws:// resolved by using WSS

### Secondary (MEDIUM confidence)
- [Office Add-In mkcert guide](https://kags.me.ke/posts/office-add-in-mkcert-localhost-ssl-certificate/) - Confirmed mkcert works for Office add-in HTTPS development
- [mkcert GitHub](https://github.com/FiloSottile/mkcert) - mkcert adds CA to system trust store, Safari/WKWebView respects it
- Microsoft Learn: [Unified manifest overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) - Unified manifest is preview-only for PowerPoint, cannot sideload on Mac

### Tertiary (LOW confidence)
- Apple Developer Forums threads about WKWebView and WSS with self-signed certs - Suggest macOS 15 may tighten restrictions, but mkcert uses a trusted CA (not self-signed)

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - Office.js CDN and native WebSocket API are well-documented official approaches
- Architecture: HIGH - XML manifest structure verified against multiple official Microsoft samples and docs
- Pitfalls: HIGH - Element ordering, wef directory, restart requirement all confirmed in official docs
- WebSocket in WKWebView: MEDIUM - mkcert approach is well-supported but macOS 15 specifics are less documented

**Research date:** 2026-02-07
**Valid until:** 2026-03-07 (Office.js and manifest format are stable; 30 days is safe)
