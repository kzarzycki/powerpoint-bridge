import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import { ConnectionPool } from './bridge.ts'
import { registerTools } from './tools.ts'

function mockWs(): WebSocket {
  return { send: vi.fn(), readyState: 1 } as unknown as WebSocket
}

async function setupMcpClient(pool: ConnectionPool) {
  const server = new McpServer({ name: 'test', version: '0.0.1' })
  registerTools(
    server,
    pool,
    () => 'test-session',
    () => 1,
  )

  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  await server.connect(serverTransport)

  const client = new Client({ name: 'test-client', version: '0.0.1' })
  await client.connect(clientTransport)

  return { client, server }
}

describe('MCP Tools', () => {
  let pool: ConnectionPool

  beforeEach(() => {
    pool = new ConnectionPool(100)
  })

  it('lists all 7 tools', async () => {
    const { client } = await setupMcpClient(pool)
    const result = await client.listTools()
    const names = result.tools.map((t) => t.name).sort()
    expect(names).toEqual(['execute_officejs', 'export_slide', 'get_presentation', 'get_slide', 'get_slide_image', 'insert_slides', 'list_presentations'])
  })

  describe('list_presentations', () => {
    it('returns empty message with no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'list_presentations', arguments: {} })
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })

    it('returns connection info when presentations are connected', async () => {
      const ws = mockWs()
      pool.add('demo.pptx', {
        ws,
        ready: true,
        presentationId: 'demo.pptx',
        filePath: '/path/demo.pptx',
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'list_presentations', arguments: {} })
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed).toHaveLength(1)
      expect(parsed[0].presentationId).toBe('demo.pptx')
      expect(parsed[0].ready).toBe(true)
    })
  })

  describe('get_presentation', () => {
    it('returns error with no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'get_presentation', arguments: {} })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('get_slide_image', () => {
    it('returns image content block on success', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_slide_image',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('getImageAsBase64')
      expect(sentJson.params.code).toContain('width: 720')

      pool.handleResponse(sentJson.id, 'response', {
        base64: 'iVBORw0KGgo=',
        slideIndex: 0,
        slideId: 'slide-abc',
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; data?: string; mimeType?: string; text?: string }>

      expect(content[0].type).toBe('image')
      expect(content[0].data).toBe('iVBORw0KGgo=')
      expect(content[0].mimeType).toBe('image/png')

      expect(content[1].type).toBe('text')
      expect(content[1].text).toContain('Slide 0')
      expect(content[1].text).toContain('slide-abc')
    })

    it('passes custom width and height to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_slide_image',
        arguments: { slideIndex: 0, width: 1280, height: 720 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('width: 1280')
      expect(sentJson.params.code).toContain('height: 720')

      pool.handleResponse(sentJson.id, 'response', {
        base64: 'abc123',
        slideIndex: 0,
        slideId: 'slide-xyz',
      })

      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'get_slide_image',
        arguments: { slideIndex: 0 },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('export_slide', () => {
    it('sends exportAsBase64 code and returns result', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'export_slide',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('exportAsBase64')

      pool.handleResponse(sentJson.id, 'response', {
        base64: 'UEsDBBQ=',
        base64Length: 8,
        slideIndex: 0,
        slideId: 'slide-1',
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.exportId).toBeDefined()
      expect(parsed.slideIndex).toBe(0)
      expect(parsed.sizeBytes).toBe(8)
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'export_slide',
        arguments: { slideIndex: 0 },
      })
      expect(result.isError).toBe(true)
    })
  })

  describe('insert_slides', () => {
    it('sends insertSlidesFromBase64 code with defaults', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_slides',
        arguments: { base64: 'UEsDBBQ=' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('insertSlidesFromBase64')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 5 })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.slideCount).toBe(5)
    })

    it('passes formatting and targetSlideId options', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_slides',
        arguments: {
          base64: 'UEsDBBQ=',
          formatting: 'UseDestinationTheme',
          targetSlideId: '267#',
          sourceSlideIds: ['256#', '267#'],
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('UseDestinationTheme')
      expect(sentJson.params.code).toContain('267#')
      expect(sentJson.params.code).toContain('sourceSlideIds')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 7 })
      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'insert_slides',
        arguments: { base64: 'UEsDBBQ=' },
      })
      expect(result.isError).toBe(true)
    })
  })

  describe('execute_officejs', () => {
    it('sends code through pool and returns result', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      // Start the tool call
      const toolPromise = client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'return 42' },
      })

      // Wait a tick for the command to be sent
      await new Promise((r) => setTimeout(r, 10))

      // Extract and respond to the command
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', 42)

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toBe('42')
    })

    it('returns error on timeout', async () => {
      vi.useFakeTimers()
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'slow code' },
      })

      // Advance past timeout
      await vi.advanceTimersByTimeAsync(200)

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('timed out')
      vi.useRealTimers()
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'return 1' },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })
})
