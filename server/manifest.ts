/**
 * Substitutes port numbers in a manifest template.
 *
 * Uses a targeted `localhost:PORT` pattern to avoid corrupting non-URL content
 * (e.g. version strings or IDs that might coincidentally contain the port number).
 */
export function substituteManifestPort(content: string, defaultPort: number, targetPort: number): string {
  if (defaultPort === targetPort) return content
  return content.replaceAll(`localhost:${defaultPort}`, `localhost:${targetPort}`)
}
