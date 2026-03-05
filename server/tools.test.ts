import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import { ConnectionPool } from './bridge.ts'
import { parseSlideRange, registerTools } from './tools.ts'

vi.mock('node:fs', () => ({ readFileSync: vi.fn() }))

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

  it('lists all 8 tools', async () => {
    const { client } = await setupMcpClient(pool)
    const result = await client.listTools()
    const names = result.tools.map((t) => t.name).sort()
    expect(names).toEqual([
      'copy_slides',
      'execute_officejs',
      'get_deck_overview',
      'get_presentation',
      'get_slide',
      'get_slide_image',
      'insert_image',
      'list_presentations',
    ])
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

  describe('copy_slides', () => {
    it('exports from source and inserts into destination', async () => {
      const sourceWs = mockWs()
      const destWs = mockWs()
      pool.add('source.pptx', {
        ws: sourceWs,
        ready: true,
        presentationId: 'source.pptx',
        filePath: '/path/source.pptx',
      })
      pool.add('dest.pptx', {
        ws: destWs,
        ready: true,
        presentationId: 'dest.pptx',
        filePath: '/path/dest.pptx',
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 2,
          sourcePresentationId: 'source.pptx',
          destinationPresentationId: 'dest.pptx',
        },
      })

      // Wait for export command to be sent to source
      await new Promise((r) => setTimeout(r, 10))

      const exportJson = JSON.parse((sourceWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(exportJson.action).toBe('executeCode')
      expect(exportJson.params.code).toContain('exportAsBase64')

      // Respond with exported Base64
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'UEsDBBQ=',
        slideIndex: 2,
        slideId: 'slide-src',
      })

      // Wait for insert command to be sent to destination
      await new Promise((r) => setTimeout(r, 10))

      const insertJson = JSON.parse((destWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(insertJson.action).toBe('executeCode')
      expect(insertJson.params.code).toContain('insertSlidesFromBase64')
      expect(insertJson.params.code).toContain('UEsDBBQ=')

      // Respond with insert result
      pool.handleResponse(insertJson.id, 'response', { slideCount: 8 })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.copied.slideIndex).toBe(2)
      expect(parsed.copied.slideId).toBe('slide-src')
      expect(parsed.destination.slideCount).toBe(8)
    })

    it('passes formatting and targetSlideId options', async () => {
      const sourceWs = mockWs()
      const destWs = mockWs()
      pool.add('a.pptx', {
        ws: sourceWs,
        ready: true,
        presentationId: 'a.pptx',
        filePath: null,
      })
      pool.add('b.pptx', {
        ws: destWs,
        ready: true,
        presentationId: 'b.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'a.pptx',
          destinationPresentationId: 'b.pptx',
          formatting: 'UseDestinationTheme',
          targetSlideId: '267#',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const exportJson = JSON.parse((sourceWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'DATA',
        slideIndex: 0,
        slideId: 'slide-0',
      })

      await new Promise((r) => setTimeout(r, 10))

      const insertJson = JSON.parse((destWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(insertJson.params.code).toContain('UseDestinationTheme')
      expect(insertJson.params.code).toContain('267#')

      pool.handleResponse(insertJson.id, 'response', { slideCount: 4 })
      await toolPromise
    })

    it('returns error when source presentation not found', async () => {
      const ws = mockWs()
      pool.add('dest.pptx', {
        ws,
        ready: true,
        presentationId: 'dest.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'missing.pptx',
          destinationPresentationId: 'dest.pptx',
        },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('missing.pptx')
    })

    it('returns error when destination presentation not found', async () => {
      const ws = mockWs()
      pool.add('source.pptx', {
        ws,
        ready: true,
        presentationId: 'source.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'source.pptx',
          destinationPresentationId: 'missing.pptx',
        },
      })

      // Export succeeds
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'DATA',
        slideIndex: 0,
        slideId: 'slide-0',
      })

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('missing.pptx')
    })
  })

  describe('insert_image', () => {
    it('passes base64 data directly into the code string', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'iVBORw0KGgoAAAANSUhEUg==',
          sourceType: 'base64',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      expect(sentJson.params.code).toContain('iVBORw0KGgoAAAANSUhEUg==')
      expect(sentJson.params.code).toContain('CoercionType.Image')

      pool.handleResponse(sentJson.id, 'response', { success: true })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.success).toBe(true)
    })

    it('reads file and base64 encodes it', async () => {
      const { readFileSync } = await import('node:fs')
      vi.mocked(readFileSync).mockReturnValue(Buffer.from([0x89, 0x50, 0x4e, 0x47]))

      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: '/path/to/image.png',
          sourceType: 'file',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      expect(readFileSync).toHaveBeenCalledWith('/path/to/image.png')

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      // The base64 of [0x89, 0x50, 0x4e, 0x47] is "iVBORw=="
      expect(sentJson.params.code).toContain(Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64'))

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('fetches URL and base64 encodes it', async () => {
      const mockArrayBuffer = new Uint8Array([1, 2, 3]).buffer
      globalThis.fetch = vi.fn().mockResolvedValue({
        ok: true,
        arrayBuffer: () => Promise.resolve(mockArrayBuffer),
      })

      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'https://example.com/image.png',
          sourceType: 'url',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      expect(globalThis.fetch).toHaveBeenCalledWith('https://example.com/image.png')

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      expect(sentJson.params.code).toContain(Buffer.from(new Uint8Array([1, 2, 3])).toString('base64'))

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('wraps with goToByIdAsync when slideIndex is provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'AAAA',
          sourceType: 'base64',
          slideIndex: 2,
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      // slideIndex 2 (0-based) → goToByIdAsync(3, ...) (1-based)
      expect(sentJson.params.code).toContain('goToByIdAsync(3,')
      expect(sentJson.params.code).toContain('GoToType.Index')

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('includes positioning options when provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'BBBB',
          sourceType: 'base64',
          left: 100,
          top: 50,
          width: 400,
          height: 300,
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('imageLeft: 100')
      expect(sentJson.params.code).toContain('imageTop: 50')
      expect(sentJson.params.code).toContain('imageWidth: 400')
      expect(sentJson.params.code).toContain('imageHeight: 300')

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'AAAA',
          sourceType: 'base64',
        },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('parseSlideRange', () => {
    it('returns null for undefined input', () => {
      expect(parseSlideRange(undefined)).toBeNull()
    })

    it('returns null for empty string', () => {
      expect(parseSlideRange('')).toBeNull()
    })

    it('parses single index', () => {
      expect(parseSlideRange('5')).toEqual([5])
    })

    it('parses comma-separated indices', () => {
      expect(parseSlideRange('2,4,7')).toEqual([2, 4, 7])
    })

    it('parses a range', () => {
      expect(parseSlideRange('0-3')).toEqual([0, 1, 2, 3])
    })

    it('parses mixed ranges and indices', () => {
      expect(parseSlideRange('0-2,5,8-10')).toEqual([0, 1, 2, 5, 8, 9, 10])
    })

    it('deduplicates overlapping ranges', () => {
      expect(parseSlideRange('0-3,2-5')).toEqual([0, 1, 2, 3, 4, 5])
    })

    it('throws on invalid index', () => {
      expect(() => parseSlideRange('abc')).toThrow('Invalid slide index')
    })

    it('throws on invalid range', () => {
      expect(() => parseSlideRange('5-2')).toThrow('Invalid slide range')
    })

    it('throws on negative index', () => {
      expect(() => parseSlideRange('-1')).toThrow('Invalid slide index')
    })
  })

  describe('get_deck_overview', () => {
    it('returns interleaved image and text blocks', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_deck_overview',
        arguments: {},
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('getImageAsBase64')
      expect(sentJson.params.code).toContain('width: 480')

      pool.handleResponse(sentJson.id, 'response', {
        slideCount: 3,
        slides: [
          {
            index: 0,
            id: 'slide-0',
            shapeCount: 2,
            shapes: [
              { name: 'Title', type: 'TextBox', id: '1', text: 'Hello World' },
              { name: 'Subtitle', type: 'TextBox', id: '2', text: 'Intro' },
            ],
            imageBase64: 'img0data',
          },
          {
            index: 1,
            id: 'slide-1',
            shapeCount: 1,
            shapes: [{ name: 'Picture', type: 'Image', id: '3' }],
            imageBase64: 'img1data',
          },
          {
            index: 2,
            id: 'slide-2',
            shapeCount: 1,
            shapes: [{ name: 'Body', type: 'TextBox', id: '4', text: 'Content here' }],
            imageBase64: 'img2data',
          },
        ],
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; text?: string; data?: string; mimeType?: string }>

      // Header text
      expect(content[0].type).toBe('text')
      expect(content[0].text).toContain('3 total slides, showing 3')

      // Slide 0: image then text
      expect(content[1].type).toBe('image')
      expect(content[1].data).toBe('img0data')
      expect(content[1].mimeType).toBe('image/png')
      expect(content[2].type).toBe('text')
      expect(content[2].text).toContain('Slide 0')
      expect(content[2].text).toContain('Hello World')
      expect(content[2].text).toContain('Intro')

      // Slide 1: image then text (no text content)
      expect(content[3].type).toBe('image')
      expect(content[4].type).toBe('text')
      expect(content[4].text).toContain('Slide 1')
      expect(content[4].text).toContain('(no text content)')

      // Slide 2: image then text
      expect(content[5].type).toBe('image')
      expect(content[6].type).toBe('text')
      expect(content[6].text).toContain('Content here')
    })

    it('skips images when includeImages is false', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_deck_overview',
        arguments: { includeImages: false },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      // Should NOT contain getImageAsBase64 in the code
      expect(sentJson.params.code).not.toContain('getImageAsBase64')

      pool.handleResponse(sentJson.id, 'response', {
        slideCount: 2,
        slides: [
          {
            index: 0,
            id: 'slide-0',
            shapeCount: 1,
            shapes: [{ name: 'Title', type: 'TextBox', id: '1', text: 'Slide text' }],
          },
          { index: 1, id: 'slide-1', shapeCount: 0, shapes: [] },
        ],
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; text?: string }>

      // Should have no image blocks at all
      const imageBlocks = content.filter((c) => c.type === 'image')
      expect(imageBlocks).toHaveLength(0)

      // Should have header + 2 slide text blocks
      expect(content).toHaveLength(3)
      expect(content[1].text).toContain('Slide text')
    })

    it('passes custom imageWidth to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_deck_overview',
        arguments: { imageWidth: 960 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('width: 960')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 0, slides: [] })
      await toolPromise
    })

    it('passes slideRange indices to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'get_deck_overview',
        arguments: { slideRange: '0-2,5' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('[0,1,2,5]')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 10, slides: [] })
      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'get_deck_overview',
        arguments: {},
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
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
