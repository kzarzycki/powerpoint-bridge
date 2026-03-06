import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from './bridge.ts'

// ---------------------------------------------------------------------------
// Local copy cache: presentationId → { localPath, revision }
// ---------------------------------------------------------------------------

export const localCopyCache = new Map<string, { localPath: string; revision: number }>()

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
// Helpers
// ---------------------------------------------------------------------------

/**
 * Parse a slide range string like "0-3,5,8-10" into a sorted, deduplicated
 * array of zero-based indices: [0,1,2,3,5,8,9,10].
 * Returns null for undefined/empty input (meaning "all slides").
 */
export function parseSlideRange(range: string | undefined): number[] | null {
  if (!range) return null
  const indices = new Set<number>()
  for (const part of range.split(',')) {
    const trimmed = part.trim()
    if (!trimmed) continue
    const dashIdx = trimmed.indexOf('-', 1)
    if (dashIdx === -1) {
      const n = Number(trimmed)
      if (!Number.isInteger(n) || n < 0) throw new Error(`Invalid slide index: "${trimmed}"`)
      indices.add(n)
    } else {
      const start = Number(trimmed.slice(0, dashIdx))
      const end = Number(trimmed.slice(dashIdx + 1))
      if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start) {
        throw new Error(`Invalid slide range: "${trimmed}"`)
      }
      for (let i = start; i <= end; i++) indices.add(i)
    }
  }
  if (indices.size === 0) return null
  return [...indices].sort((a, b) => a - b)
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
        .describe(
          'Image width in pixels. Default: 720. Height auto-calculated to preserve aspect ratio unless also specified.',
        ),
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
        const description = `Slide ${result.slideIndex} (ID: ${result.slideId})${warning ?? ''}`

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

  // --- Tool: copy_slides ---
  server.tool(
    'copy_slides',
    'Copies slides between two open presentations entirely server-side — the Base64 data never enters Claude context. Exports from the source presentation and inserts into the destination in a single operation. Both presentations must be connected to the bridge. Requires PowerPointApi 1.8.',
    {
      sourceSlideIndex: z.number().int().min(0).describe('Zero-based slide index to copy from the source presentation'),
      sourcePresentationId: z.string().describe('Source presentation ID from list_presentations'),
      destinationPresentationId: z.string().describe('Destination presentation ID from list_presentations'),
      targetSlideId: z
        .string()
        .optional()
        .describe(
          'Insert after this slide ID in destination (format: "nnn#" or "#mmmmmmmmm" or "nnn#mmmmmmmmm"). If omitted, inserts at the beginning.',
        ),
      formatting: z
        .enum(['KeepSourceFormatting', 'UseDestinationTheme'])
        .optional()
        .describe('Formatting mode. Default: KeepSourceFormatting.'),
    },
    async ({ sourceSlideIndex, sourcePresentationId, destinationPresentationId, targetSlideId, formatting }) => {
      try {
        // Step 1: Export slide from source presentation
        const source = pool.resolveTarget(sourcePresentationId)
        const exportCode = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${sourceSlideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${sourceSlideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${sourceSlideIndex}];
          var result = slide.exportAsBase64();
          await context.sync();
          return { base64: result.value, slideIndex: ${sourceSlideIndex}, slideId: slide.id };
        `
        const exported = (await pool.sendCommand('executeCode', { code: exportCode }, source.ws)) as {
          base64: string
          slideIndex: number
          slideId: string
        }

        // Step 2: Insert into destination presentation
        const dest = pool.resolveTarget(destinationPresentationId)
        const optionsParts: string[] = []
        if (formatting) {
          optionsParts.push(`formatting: "${formatting}"`)
        }
        if (targetSlideId) {
          optionsParts.push(`targetSlideId: "${targetSlideId}"`)
        }
        const optionsArg = optionsParts.length > 0 ? `, { ${optionsParts.join(', ')} }` : ''

        const insertCode = `
          context.presentation.insertSlidesFromBase64("${exported.base64}"${optionsArg});
          await context.sync();
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return { slideCount: slides.items.length };
        `
        const inserted = (await pool.sendCommand('executeCode', { code: insertCode }, dest.ws)) as {
          slideCount: number
        }

        const warning = getConcurrentWarning(getSessionId(), dest.presentationId, getActiveSessionCount())
        const text =
          JSON.stringify(
            {
              copied: { slideIndex: exported.slideIndex, slideId: exported.slideId },
              destination: { slideCount: inserted.slideCount },
            },
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

  // --- Tool: insert_image ---
  server.tool(
    'insert_image',
    'Inserts an image onto a slide using Office.js setSelectedDataAsync with CoercionType.Image. Accepts a file path, URL, or raw base64 data. Optionally navigate to a specific slide first and control position/size in points.',
    {
      source: z.string().describe('File path, URL, or base64 image data depending on sourceType'),
      sourceType: z
        .enum(['file', 'url', 'base64'])
        .describe(
          'How to interpret source: "file" reads from disk, "url" fetches from network, "base64" uses data directly',
        ),
      slideIndex: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          'Zero-based slide index to navigate to before inserting. If omitted, inserts on the currently active slide.',
        ),
      left: z.number().optional().describe('Horizontal position in points (1 point = 1/72 inch)'),
      top: z.number().optional().describe('Vertical position in points'),
      width: z.number().optional().describe('Image width in points'),
      height: z.number().optional().describe('Image height in points'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ source, sourceType, slideIndex, left, top, width, height, presentationId }) => {
      try {
        // Step 1: Resolve image to base64
        let base64Data: string
        if (sourceType === 'file') {
          base64Data = readFileSync(source).toString('base64')
        } else if (sourceType === 'url') {
          const resp = await fetch(source)
          if (!resp.ok) {
            throw new Error(`Failed to fetch image from URL: ${resp.status} ${resp.statusText}`)
          }
          const buf = await resp.arrayBuffer()
          base64Data = Buffer.from(buf).toString('base64')
        } else {
          base64Data = source
        }

        // Step 2: Build options object string with only provided params
        const optionsParts: string[] = ['coercionType: Office.CoercionType.Image']
        if (left !== undefined) optionsParts.push(`imageLeft: ${left}`)
        if (top !== undefined) optionsParts.push(`imageTop: ${top}`)
        if (width !== undefined) optionsParts.push(`imageWidth: ${width}`)
        if (height !== undefined) optionsParts.push(`imageHeight: ${height}`)
        const optionsStr = `{ ${optionsParts.join(', ')} }`

        // Step 3: Build the setSelectedDataAsync call
        const insertCall = `Office.context.document.setSelectedDataAsync("${base64Data}", ${optionsStr}, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ success: true });
        } else {
          reject(new Error(result.error.message));
        }
      });`

        // Step 4: Wrap with goToByIdAsync if slideIndex is provided
        let code: string
        if (slideIndex !== undefined) {
          code = `return new Promise(function(resolve, reject) {
      Office.context.document.goToByIdAsync(${slideIndex + 1}, Office.GoToType.Index, function(navResult) {
        if (navResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Navigation failed: " + navResult.error.message));
          return;
        }
        ${insertCall}
      });
    });`
        } else {
          code = `return new Promise(function(resolve, reject) {
      ${insertCall}
    });`
        }

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

  // --- Tool: get_deck_overview ---
  server.tool(
    'get_deck_overview',
    'Returns a visual overview of all (or selected) slides in one call — thumbnails interleaved with text metadata. Much more efficient than calling get_slide + get_slide_image per slide. Use this to review or audit an entire presentation quickly.',
    {
      slideRange: z
        .string()
        .optional()
        .describe('Slide indices to include, e.g. "0-5", "2,4,7", "0-2,5,8-10". Omit for all slides.'),
      imageWidth: z
        .number()
        .int()
        .min(120)
        .max(1920)
        .optional()
        .describe('Thumbnail width in pixels. Default: 480. Height auto-calculated to preserve aspect ratio.'),
      includeImages: z
        .boolean()
        .optional()
        .describe('Include slide thumbnails. Default: true. Set false for text-only overview (faster).'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideRange, imageWidth, includeImages, presentationId }) => {
      try {
        const indices = parseSlideRange(slideRange)
        const width = imageWidth ?? 480
        const withImages = includeImages !== false

        // Build the indices array literal for Office.js, or null for "all"
        const indicesJs = indices ? JSON.stringify(indices) : 'null'

        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          var requestedIndices = ${indicesJs};
          var indicesToProcess = requestedIndices || [];
          if (!requestedIndices) {
            for (var i = 0; i < slides.items.length; i++) indicesToProcess.push(i);
          }
          // Validate indices
          for (var i = 0; i < indicesToProcess.length; i++) {
            if (indicesToProcess[i] >= slides.items.length) {
              throw new Error("Slide index " + indicesToProcess[i] + " out of range (presentation has " + slides.items.length + " slides)");
            }
          }
          // Load shapes for all requested slides
          for (var i = 0; i < indicesToProcess.length; i++) {
            slides.items[indicesToProcess[i]].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < indicesToProcess.length; i++) {
            var idx = indicesToProcess[i];
            var slide = slides.items[idx];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              var info = { name: s.name, type: s.type, id: s.id };
              try {
                s.textFrame.load("textRange");
                await context.sync();
                info.text = s.textFrame.textRange.text;
              } catch (e) {}
              shapes.push(info);
            }
            var slideData = { index: idx, id: slide.id, shapeCount: shapes.length, shapes: shapes };
            ${
              withImages
                ? `var img = slide.getImageAsBase64({ width: ${width} });
            await context.sync();
            slideData.imageBase64 = img.value;`
                : ''
            }
            output.push(slideData);
          }
          return { slideCount: slides.items.length, slides: output };
        `
        const target = pool.resolveTarget(presentationId)
        const result = (await pool.sendCommand('executeCode', { code }, target.ws, 120_000)) as {
          slideCount: number
          slides: Array<{
            index: number
            id: string
            shapeCount: number
            shapes: Array<{ name: string; type: string; id: string; text?: string }>
            imageBase64?: string
          }>
        }
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())

        // Build interleaved content blocks
        const content: Array<{ type: 'text'; text: string } | { type: 'image'; data: string; mimeType: string }> = []
        const showing = result.slides.length
        const header = `Deck overview: ${result.slideCount} total slides, showing ${showing}${warning ?? ''}`
        content.push({ type: 'text' as const, text: header })

        for (const slide of result.slides) {
          if (slide.imageBase64) {
            content.push({ type: 'image' as const, data: slide.imageBase64, mimeType: 'image/png' })
          }
          const textParts = slide.shapes.filter((s) => s.text).map((s) => s.text!)
          const shapeText = textParts.length > 0 ? `\n${textParts.join('\n')}` : '\n(no text content)'
          content.push({
            type: 'text' as const,
            text: `--- Slide ${slide.index} | ${slide.shapeCount} shapes ---${shapeText}`,
          })
        }

        return { content }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: get_local_copy ---
  server.tool(
    'get_local_copy',
    'Returns a local file path for the presentation. For local files, returns the existing path. For SharePoint/cloud files, exports server-side and saves to a temp .pptx. Caches by revision number — re-exports only when the presentation has been saved since last export.',
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const filePath = target.filePath

        // Local file — already on disk
        if (filePath && !filePath.startsWith('http')) {
          if (!existsSync(filePath)) {
            return { content: [{ type: 'text' as const, text: `Error: Local file not found: ${filePath}` }], isError: true }
          }
          return { content: [{ type: 'text' as const, text: JSON.stringify({ localPath: filePath, source: 'local' }) }] }
        }

        // Cloud file — check revision for cache validity
        const revCode = `
          var p = context.presentation.properties;
          p.load("revisionNumber");
          await context.sync();
          return p.revisionNumber;
        `
        const currentRevision = await pool.sendCommand('executeCode', { code: revCode }, target.ws) as number

        const cached = localCopyCache.get(target.presentationId)
        if (cached && cached.revision === currentRevision && existsSync(cached.localPath)) {
          return { content: [{ type: 'text' as const, text: JSON.stringify({ localPath: cached.localPath, source: 'cached', revision: currentRevision }) }] }
        }

        // Export fresh copy
        const exportCode = `
          var r = context.presentation.exportAsBase64();
          await context.sync();
          return r.value;
        `
        const base64 = await pool.sendCommand('executeCode', { code: exportCode }, target.ws, 120_000) as string

        const filename = filePath
          ? decodeURIComponent(filePath.split('/').pop() || 'presentation.pptx')
          : 'presentation.pptx'
        const dest = join(tmpdir(), `pptbridge-${Date.now()}-${filename}`)
        writeFileSync(dest, Buffer.from(base64, 'base64'))

        localCopyCache.set(target.presentationId, { localPath: dest, revision: currentRevision })

        return { content: [{ type: 'text' as const, text: JSON.stringify({ localPath: dest, source: 'exported', revision: currentRevision }) }] }
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
