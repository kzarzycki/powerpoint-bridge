import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import type { AddinConnection } from './bridge.ts'
import { ConnectionPool } from './bridge.ts'

function mockWs(overrides: Partial<WebSocket> = {}): WebSocket {
  return {
    send: vi.fn(),
    readyState: 1,
    ...overrides,
  } as unknown as WebSocket
}

function makeConn(ws: WebSocket, opts: Partial<AddinConnection> = {}): AddinConnection {
  return {
    ws,
    ready: true,
    presentationId: opts.presentationId ?? 'test.pptx',
    filePath: opts.filePath ?? '/path/test.pptx',
    ...opts,
  }
}

describe('ConnectionPool', () => {
  let pool: ConnectionPool

  beforeEach(() => {
    pool = new ConnectionPool(100) // 100ms timeout for fast tests
  })

  describe('resolveTarget', () => {
    it('throws when no connections exist', () => {
      expect(() => pool.resolveTarget()).toThrow('No presentations connected')
    })

    it('returns the single connection when only one exists', () => {
      const ws = mockWs()
      const conn = makeConn(ws)
      pool.add('test.pptx', conn)
      expect(pool.resolveTarget()).toBe(conn)
    })

    it('throws when single connection is not ready', () => {
      const ws = mockWs()
      pool.add('test.pptx', makeConn(ws, { ready: false }))
      expect(() => pool.resolveTarget()).toThrow('not ready')
    })

    it('throws listing available IDs when multiple connections exist without ID', () => {
      pool.add('a.pptx', makeConn(mockWs(), { presentationId: 'a.pptx' }))
      pool.add('b.pptx', makeConn(mockWs(), { presentationId: 'b.pptx' }))
      expect(() => pool.resolveTarget()).toThrow('Multiple presentations connected')
      expect(() => pool.resolveTarget()).toThrow('a.pptx')
      expect(() => pool.resolveTarget()).toThrow('b.pptx')
    })

    it('returns correct connection when presentationId is specified', () => {
      const wsA = mockWs()
      const wsB = mockWs()
      const connA = makeConn(wsA, { presentationId: 'a.pptx' })
      const connB = makeConn(wsB, { presentationId: 'b.pptx' })
      pool.add('a.pptx', connA)
      pool.add('b.pptx', connB)
      expect(pool.resolveTarget('b.pptx')).toBe(connB)
    })

    it('throws when specified presentationId is not found', () => {
      pool.add('a.pptx', makeConn(mockWs()))
      expect(() => pool.resolveTarget('missing.pptx')).toThrow('Presentation not found: missing.pptx')
    })

    it('throws when specified presentation is not ready', () => {
      pool.add('a.pptx', makeConn(mockWs(), { ready: false, presentationId: 'a.pptx' }))
      expect(() => pool.resolveTarget('a.pptx')).toThrow('not ready')
    })
  })

  describe('removeBySocket', () => {
    it('removes the connection and returns its ID', () => {
      const ws = mockWs()
      pool.add('test.pptx', makeConn(ws))
      expect(pool.size).toBe(1)
      const removed = pool.removeBySocket(ws)
      expect(removed).toBe('test.pptx')
      expect(pool.size).toBe(0)
    })

    it('returns null when socket is not found', () => {
      expect(pool.removeBySocket(mockWs())).toBeNull()
    })
  })

  describe('generateId', () => {
    it('returns documentUrl when provided', () => {
      expect(pool.generateId('/path/to/file.pptx')).toBe('/path/to/file.pptx')
    })

    it('generates incrementing untitled IDs when no URL', () => {
      expect(pool.generateId(null)).toBe('untitled-1')
      expect(pool.generateId(null)).toBe('untitled-2')
    })
  })

  describe('sendCommand', () => {
    it('sends JSON command and resolves when response arrives', async () => {
      const ws = mockWs()
      const promise = pool.sendCommand('executeCode', { code: 'test' }, ws)

      // Extract the command ID from what was sent
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.type).toBe('command')
      expect(sentJson.action).toBe('executeCode')

      // Simulate response
      pool.handleResponse(sentJson.id, 'response', { result: 42 })
      await expect(promise).resolves.toEqual({ result: 42 })
    })

    it('rejects when error response arrives', async () => {
      const ws = mockWs()
      const promise = pool.sendCommand('executeCode', { code: 'bad' }, ws)

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'error', undefined, 'Something went wrong')

      await expect(promise).rejects.toThrow('Something went wrong')
    })

    it('rejects on timeout', async () => {
      vi.useFakeTimers()
      const ws = mockWs()
      const promise = pool.sendCommand('executeCode', { code: 'slow' }, ws)

      vi.advanceTimersByTime(200) // past the 100ms timeout
      await expect(promise).rejects.toThrow('Command timed out')
      vi.useRealTimers()
    })
  })

  describe('rejectPendingForSocket', () => {
    it('rejects only pending requests for the given socket', async () => {
      const wsA = mockWs()
      const wsB = mockWs()

      const promiseA = pool.sendCommand('executeCode', { code: 'a' }, wsA)
      const promiseB = pool.sendCommand('executeCode', { code: 'b' }, wsB)

      pool.rejectPendingForSocket(wsA)

      await expect(promiseA).rejects.toThrow('Add-in disconnected')

      // promiseB should still be pending — resolve it manually
      const sentB = JSON.parse((wsB.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentB.id, 'response', 'ok')
      await expect(promiseB).resolves.toBe('ok')
    })
  })
})
