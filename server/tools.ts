import { randomUUID } from 'node:crypto'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from './bridge.ts'

// ---------------------------------------------------------------------------
// Server-side export store (keeps large Base64 data out of Claude's context)
// ---------------------------------------------------------------------------

const exportStore = new Map<string, { base64: string; createdAt: number }>()
const EXPORT_TTL_MS = 5 * 60 * 1000 // 5 minutes

function storeExport(base64: string): string {
  // Clean expired entries
  const now = Date.now()
  for (const [id, entry] of exportStore) {
    if (now - entry.createdAt > EXPORT_TTL_MS) exportStore.delete(id)
  }
  const id = randomUUID().slice(0, 8)
  exportStore.set(id, { base64, createdAt: now })
  return id
}

function getExport(id: string): string | null {
  const entry = exportStore.get(id)
  if (!entry) return null
  if (Date.now() - entry.createdAt > EXPORT_TTL_MS) {
    exportStore.delete(id)
    return null
  }
  return entry.base64
}

// ---------------------------------------------------------------------------
// Concurrent access warning helper
// ---------------------------------------------------------------------------

const sessionConcurrentWarnings = new Map<string, Set<string>>()

export function getConcurrentWarning(
  mcpSessionId: string | undefined,
  presentationId: string,
  activeSessions: number,
): string | null {
  if (!mcpSessionId) return null
  if (activeSessions <= 1) return null

  const warned = sessionConcurrentWarnings.get(mcpSessionId)
  if (warned?.has(presentationId)) return null

  if (!warned) {
    sessionConcurrentWarnings.set(mcpSessionId, new Set([presentationId]))
  } else {
    warned.add(presentationId)
  }

  return '\n\nNote: Other MCP sessions are also connected to the bridge. If they target this presentation, changes apply immediately (last-write-wins).'
}

export function clearSessionWarnings(sessionId: string): void {
  sessionConcurrentWarnings.delete(sessionId)
}

// ---------------------------------------------------------------------------
// Tool registration
// ---------------------------------------------------------------------------

export function registerTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  // --- Tool: list_presentations ---
  server.tool(
    'list_presentations',
    'Lists all PowerPoint presentations currently connected to the bridge server. Shows presentation IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the presentationId to pass to other tools when multiple presentations are open.',
    async () => {
      const presentations = []
      for (const [id, conn] of pool.entries()) {
        presentations.push({
          presentationId: id,
          filePath: conn.filePath,
          ready: conn.ready,
        })
      }
      return {
        content: [
          {
            type: 'text' as const,
            text:
              presentations.length === 0
                ? 'No presentations connected. Open a PowerPoint file with the bridge add-in loaded.'
                : JSON.stringify(presentations, null, 2),
          },
        ],
      }
    },
  )

  // --- Tool: get_presentation ---
  server.tool(
    'get_presentation',
    "Returns the structure of the currently open PowerPoint presentation including all slides with their IDs and shape summaries (count, names, types). Use this first to understand what's in the presentation before making changes.",
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ presentationId }) => {
      try {
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          for (var i = 0; i < slides.items.length; i++) {
            slides.items[i].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              shapes.push({ name: s.name, type: s.type, id: s.id });
            }
            output.push({ index: i, id: slide.id, shapeCount: shapes.length, shapes: shapes });
          }
          return output;
        `
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: get_slide ---
  server.tool(
    'get_slide',
    'Returns detailed information about all shapes on a specific slide, including text content, positions (left, top in points), sizes (width, height in points), and fill colors. Use slideIndex from get_presentation results (zero-based).',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from get_presentation results'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, presentationId }) => {
      try {
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          slide.shapes.load("items");
          await context.sync();
          var shapes = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var info = {
              name: s.name,
              type: s.type,
              id: s.id,
              left: s.left,
              top: s.top,
              width: s.width,
              height: s.height
            };
            try {
              s.textFrame.load("textRange");
              await context.sync();
              info.text = s.textFrame.textRange.text;
            } catch (e) {
              // Shape has no text frame (e.g., images, connectors)
            }
            try {
              s.fill.load("foregroundColor,type");
              await context.sync();
              info.fill = { type: s.fill.type, color: s.fill.foregroundColor };
            } catch (e) {
              // Shape has no fill or fill not accessible
            }
            shapes.push(info);
          }
          return { slideIndex: ${slideIndex}, slideId: slide.id, shapes: shapes };
        `
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: get_slide_image ---
  server.tool(
    'get_slide_image',
    'Captures a visual screenshot of a specific slide as a PNG image. Use this to SEE what a slide looks like — useful for verifying layout after changes or understanding content visually. Requires PowerPoint 16.96+ (PowerPointApi 1.8).',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from get_presentation results'),
      width: z
        .number()
        .int()
        .min(1)
        .max(4096)
        .optional()
        .describe('Image width in pixels. Default: 720. Height auto-calculated to preserve aspect ratio unless also specified.'),
      height: z
        .number()
        .int()
        .min(1)
        .max(4096)
        .optional()
        .describe('Image height in pixels. If omitted, auto-calculated from width to preserve aspect ratio.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, width, height, presentationId }) => {
      try {
        const imgWidth = width ?? 720
        const optionsParts: string[] = [`width: ${imgWidth}`]
        if (height !== undefined) {
          optionsParts.push(`height: ${height}`)
        }
        const optionsStr = `{ ${optionsParts.join(', ')} }`

        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          var result = slide.getImageAsBase64(${optionsStr});
          await context.sync();
          return { base64: result.value, slideIndex: ${slideIndex}, slideId: slide.id };
        `
        const target = pool.resolveTarget(presentationId)
        const result = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
          base64: string
          slideIndex: number
          slideId: string
        }
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const description = `Slide ${result.slideIndex} (ID: ${result.slideId})` + (warning ?? '')

        return {
          content: [
            {
              type: 'image' as const,
              data: result.base64,
              mimeType: 'image/png',
            },
            {
              type: 'text' as const,
              text: description,
            },
          ],
        }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        const hint =
          message.includes('getImageAsBase64') || message.includes('not a function')
            ? ' (This API requires PowerPoint 16.96+ with PowerPointApi 1.8 support)'
            : ''
        return { content: [{ type: 'text' as const, text: `Error: ${message}${hint}` }], isError: true }
      }
    },
  )

  // --- Tool: export_slide ---
  server.tool(
    'export_slide',
    'Exports a slide as a standalone .pptx stored server-side. Returns an exportId reference that can be passed to insert_slides to copy slides between presentations. The Base64 data stays on the server (never sent to Claude). Export expires after 5 minutes. Requires PowerPointApi 1.8.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index to export'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, presentationId }) => {
      try {
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          var result = slide.exportAsBase64();
          await context.sync();
          return { base64: result.value, slideIndex: ${slideIndex}, slideId: slide.id };
        `
        const target = pool.resolveTarget(presentationId)
        const result = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
          base64: string
          slideIndex: number
          slideId: string
        }
        const exportId = storeExport(result.base64)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text =
          JSON.stringify(
            { exportId, slideIndex: result.slideIndex, slideId: result.slideId, sizeBytes: result.base64.length },
            null,
            2,
          ) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: insert_slides ---
  server.tool(
    'insert_slides',
    'Inserts slides from a Base64-encoded .pptx into the current presentation. Use with export_slide to copy slides between open presentations. Options: choose source slides by ID, target position, and whether to keep source formatting or use destination theme.',
    {
      base64: z.string().describe('Base64-encoded .pptx data (from export_slide or any .pptx file)'),
      targetSlideId: z
        .string()
        .optional()
        .describe('Insert after this slide ID (format: "nnn#" or "#mmmmmmmmm" or "nnn#mmmmmmmmm"). If omitted, inserts at the beginning.'),
      sourceSlideIds: z
        .array(z.string())
        .optional()
        .describe('Array of source slide IDs to insert (same format as targetSlideId). If omitted, inserts all slides.'),
      formatting: z
        .enum(['KeepSourceFormatting', 'UseDestinationTheme'])
        .optional()
        .describe('Formatting mode. Default: KeepSourceFormatting.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ base64, targetSlideId, sourceSlideIds, formatting, presentationId }) => {
      try {
        const optionsParts: string[] = []
        if (formatting) {
          optionsParts.push(`formatting: "${formatting}"`)
        }
        if (targetSlideId) {
          optionsParts.push(`targetSlideId: "${targetSlideId}"`)
        }
        if (sourceSlideIds && sourceSlideIds.length > 0) {
          optionsParts.push(`sourceSlideIds: ${JSON.stringify(sourceSlideIds)}`)
        }
        const optionsArg = optionsParts.length > 0 ? `, { ${optionsParts.join(', ')} }` : ''

        const code = `
          context.presentation.insertSlidesFromBase64("${base64}"${optionsArg});
          await context.sync();
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return { slideCount: slides.items.length };
        `
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: execute_officejs ---
  server.tool(
    'execute_officejs',
    "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. For positioning, all values are in points (1 point = 1/72 inch). Common operations: add shapes, set text, change colors, add/delete slides.",
    {
      code: z
        .string()
        .describe(
          "Office.js code to execute. Runs inside PowerPoint.run() with 'context' available. Use 'return' to send back a result.",
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ code, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )
}
