#!/usr/bin/env node

/**
 * Sideloads the add-in manifest into PowerPoint's WEF directory,
 * substituting the correct port from BRIDGE_PORT env var.
 *
 * Usage:
 *   node scripts/sideload.mjs           # HTTP manifest, default port 8080
 *   node scripts/sideload.mjs --tls     # HTTPS manifest, default port 8443
 *   BRIDGE_PORT=9090 node scripts/sideload.mjs  # Custom port
 */

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { homedir } from 'node:os'
import { fileURLToPath } from 'node:url'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const PROJECT_ROOT = resolve(SCRIPT_DIR, '..')

const tls = process.argv.includes('--tls')
const defaultPort = tls ? 8443 : 8080
const port = Number(process.env.BRIDGE_PORT) || defaultPort

const manifestName = tls ? 'manifest-https.xml' : 'manifest.xml'
const src = resolve(PROJECT_ROOT, 'addin', manifestName)
const wefDir = resolve(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef')
const dest = resolve(wefDir, 'manifest.xml')

// Read template and substitute port (mirrors server/manifest.ts substituteManifestPort)
let content = readFileSync(src, 'utf8')
if (port !== defaultPort) {
  content = content.replaceAll(`localhost:${defaultPort}`, `localhost:${port}`)
}

// Write to WEF directory
mkdirSync(wefDir, { recursive: true })
writeFileSync(dest, content)

// Write version marker
const pkg = JSON.parse(readFileSync(resolve(PROJECT_ROOT, 'package.json'), 'utf8'))
writeFileSync(resolve(PROJECT_ROOT, '.sideloaded'), `${pkg.version}:${port}`)

console.log(`[sideload] Manifest installed (${tls ? 'HTTPS' : 'HTTP'}, port ${port})`)
