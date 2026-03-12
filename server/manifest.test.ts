import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { substituteManifestPort } from './manifest.ts'

describe('substituteManifestPort', () => {
  it('returns content unchanged when ports match', () => {
    const content = 'http://localhost:8080/index.html'
    expect(substituteManifestPort(content, 8080, 8080)).toBe(content)
  })

  it('replaces all localhost:PORT occurrences', () => {
    const content = ['http://localhost:8080/a', 'http://localhost:8080/b', 'http://localhost:8080/c'].join('\n')
    const result = substituteManifestPort(content, 8080, 9090)
    expect(result).not.toContain('localhost:8080')
    expect(result.match(/localhost:9090/g)?.length).toBe(3)
  })

  it('does not replace port numbers outside localhost URLs', () => {
    const content = '<Version>1.8080.0</Version>\nhttp://localhost:8080/index.html'
    const result = substituteManifestPort(content, 8080, 9090)
    expect(result).toContain('<Version>1.8080.0</Version>')
    expect(result).toContain('http://localhost:9090/index.html')
  })

  it('replaces all 9 occurrences in manifest.xml', () => {
    const manifestPath = resolve(__dirname, '..', 'addin', 'manifest.xml')
    const content = readFileSync(manifestPath, 'utf8')
    const result = substituteManifestPort(content, 8080, 9090)
    expect(result).not.toContain('localhost:8080')
    expect(result.match(/localhost:9090/g)?.length).toBe(9)
  })

  it('replaces all 9 occurrences in manifest-https.xml', () => {
    const manifestPath = resolve(__dirname, '..', 'addin', 'manifest-https.xml')
    const content = readFileSync(manifestPath, 'utf8')
    const result = substituteManifestPort(content, 8443, 9443)
    expect(result).not.toContain('localhost:8443')
    expect(result.match(/localhost:9443/g)?.length).toBe(9)
  })
})
