import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import JSZip from 'jszip'
import type { WebSocket } from 'ws'
import type { ConnectionPool } from './bridge.ts'

const SLIDE_XML_PATH = 'ppt/slides/slide1.xml'

// ---------------------------------------------------------------------------
// Export slide as base64 zip via add-in
// ---------------------------------------------------------------------------

export interface ExportedSlide {
  base64: string
  slideId: string
  prevSlideId: string | null
}

export async function exportSlide(
  pool: ConnectionPool,
  slideIndex: number,
  targetWs: WebSocket,
  timeout?: number,
): Promise<ExportedSlide> {
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
    var prevSlideId = ${slideIndex} > 0 ? slides.items[${slideIndex} - 1].id : null;
    return { base64: result.value, slideId: slide.id, prevSlideId: prevSlideId };
  `
  return (await pool.sendCommand('executeCode', { code }, targetWs, timeout)) as ExportedSlide
}

// ---------------------------------------------------------------------------
// Reimport modified slide (delete old + insert modified at same position)
// ---------------------------------------------------------------------------

export async function reimportSlide(
  pool: ConnectionPool,
  modifiedBase64: string,
  slideId: string,
  prevSlideId: string | null,
  targetWs: WebSocket,
  timeout?: number,
): Promise<void> {
  const optionsParts = ['formatting: "KeepSourceFormatting"']
  if (prevSlideId) {
    optionsParts.push(`targetSlideId: "${prevSlideId}"`)
  }
  const optionsStr = `{ ${optionsParts.join(', ')} }`

  // The base64 is embedded directly in the code string — it stays server-side
  // and never enters Claude's context.
  const code = `
    var slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    // Find and delete the original slide by ID
    var found = false;
    for (var i = 0; i < slides.items.length; i++) {
      if (slides.items[i].id === "${slideId}") {
        slides.items[i].delete();
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error("Original slide not found for reimport (ID: ${slideId})");
    }
    await context.sync();
    // Insert modified slide at the correct position
    context.presentation.insertSlidesFromBase64("${modifiedBase64}", ${optionsStr});
    await context.sync();
    return { success: true };
  `
  await pool.sendCommand('executeCode', { code }, targetWs, timeout)
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

export function parseSlideXml(xmlString: string): Document {
  return new DOMParser().parseFromString(xmlString, 'text/xml')
}

export function serializeXml(doc: Document | Element): string {
  return new XMLSerializer().serializeToString(doc)
}

/**
 * Find a <p:sp> shape element by matching <p:cNvPr id="shapeId">.
 * Shape IDs from Office.js are strings like "5" — the OOXML id attribute matches.
 */
export function findShapeById(doc: Document, shapeId: string): Element | null {
  const shapes = doc.getElementsByTagNameNS(NS_P, 'sp')
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i]!
    const nvSpPr = shape.getElementsByTagNameNS(NS_P, 'nvSpPr')
    if (nvSpPr.length === 0) continue
    const cNvPr = nvSpPr[0]!.getElementsByTagNameNS(NS_P, 'cNvPr')
    if (cNvPr.length > 0 && cNvPr[0]!.getAttribute('id') === shapeId) {
      return shape
    }
  }
  return null
}

/**
 * Extract all <a:p> elements from a shape's <p:txBody> as serialized XML string.
 */
export function extractParagraphs(shape: Element): string {
  const txBody = shape.getElementsByTagNameNS(NS_P, 'txBody')
  if (txBody.length === 0) {
    throw new Error('Shape has no text body (<p:txBody>)')
  }
  const body = txBody[0]!
  const paragraphs = body.getElementsByTagNameNS(NS_A, 'p')
  const parts: string[] = []
  for (let i = 0; i < paragraphs.length; i++) {
    parts.push(serializeXml(paragraphs[i]!))
  }
  return parts.join('')
}

/**
 * Replace all <a:p> elements in a shape's <p:txBody>, preserving <a:bodyPr> and <a:lstStyle>.
 */
export function replaceParagraphs(doc: Document, shape: Element, paragraphXml: string): void {
  const txBody = shape.getElementsByTagNameNS(NS_P, 'txBody')
  if (txBody.length === 0) {
    throw new Error('Shape has no text body (<p:txBody>)')
  }
  const body = txBody[0]!

  // Save <a:bodyPr> and <a:lstStyle>
  const bodyPr = body.getElementsByTagNameNS(NS_A, 'bodyPr')[0] ?? null
  const lstStyle = body.getElementsByTagNameNS(NS_A, 'lstStyle')[0] ?? null

  // Remove all children
  while (body.firstChild) {
    body.removeChild(body.firstChild)
  }

  // Re-append preserved elements
  if (bodyPr) body.appendChild(bodyPr)
  if (lstStyle) body.appendChild(lstStyle)

  // Parse and import new paragraphs
  const wrapper = `<wrapper xmlns:a="${NS_A}">${paragraphXml}</wrapper>`
  const fragDoc = new DOMParser().parseFromString(wrapper, 'text/xml')
  const newParagraphs = fragDoc.documentElement.childNodes
  for (let i = 0; i < newParagraphs.length; i++) {
    const imported = doc.importNode(newParagraphs[i]!, true)
    body.appendChild(imported)
  }
}

/**
 * Replace a shape's <p:sp> element in the document with new XML.
 */
export function replaceShape(doc: Document, oldShape: Element, newShapeXml: string): void {
  const fragDoc = new DOMParser().parseFromString(newShapeXml, 'text/xml')
  const imported = doc.importNode(fragDoc.documentElement, true)
  oldShape.parentNode!.replaceChild(imported, oldShape)
}

// ---------------------------------------------------------------------------
// Zip helpers
// ---------------------------------------------------------------------------

/** List all file paths in a zip (excludes directory entries). */
export function listZipPaths(zip: JSZip): string[] {
  const paths: string[] = []
  zip.forEach((relativePath, file) => {
    if (!file.dir) paths.push(relativePath)
  })
  return paths.sort()
}

/** Extract specific files from a base64 zip. If paths omitted, extracts all text/xml files. */
export async function extractZipFiles(
  base64: string,
  paths?: string[],
): Promise<{ zip: JSZip; files: Record<string, string> }> {
  const zip = await JSZip.loadAsync(base64, { base64: true })
  const files: Record<string, string> = {}

  if (paths) {
    for (const path of paths) {
      const file = zip.file(path)
      if (!file) {
        throw new Error(`File not found in zip: ${path}`)
      }
      files[path] = await file.async('string')
    }
  } else {
    // Auto-discover: extract all text/xml files (skip binary media)
    const allPaths = listZipPaths(zip)
    for (const path of allPaths) {
      if (path.endsWith('/')) continue // skip directories
      if (path.match(/\.(xml|rels)$/) || path === '[Content_Types].xml') {
        const file = zip.file(path)
        if (file) {
          files[path] = await file.async('string')
        }
      }
    }
  }

  return { zip, files }
}

/** Update multiple files in a zip and regenerate as base64. Can add new files. */
export async function updateZipFiles(zip: JSZip, files: Record<string, string>): Promise<string> {
  for (const [path, content] of Object.entries(files)) {
    zip.file(path, content)
  }
  return await zip.generateAsync({ type: 'base64' })
}

const CONTENT_TYPE_MAP: Record<string, string> = {
  'ppt/charts/': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
}

/**
 * Auto-register Content_Types for newly added files.
 * Reads [Content_Types].xml, adds <Override> entries for known file types,
 * and writes back. Only adds entries not already present.
 */
export async function autoRegisterContentTypes(zip: JSZip, newPaths: string[]): Promise<void> {
  const overrides: Array<{ partName: string; contentType: string }> = []
  for (const path of newPaths) {
    for (const [prefix, contentType] of Object.entries(CONTENT_TYPE_MAP)) {
      if (path.startsWith(prefix) && path.endsWith('.xml')) {
        overrides.push({ partName: `/${path}`, contentType })
      }
    }
  }
  if (overrides.length === 0) return

  const ctFile = zip.file('[Content_Types].xml')
  if (!ctFile) return

  let ctXml = await ctFile.async('string')
  for (const { partName, contentType } of overrides) {
    if (ctXml.includes(`PartName="${partName}"`)) continue
    // Insert before closing </Types>
    const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`
    ctXml = ctXml.replace('</Types>', `${override}</Types>`)
  }
  zip.file('[Content_Types].xml', ctXml)
}

// Legacy wrappers — used by existing read/edit_slide_text and read/edit_slide_xml tools
export async function extractSlideXmlFromZip(base64: string): Promise<{ zip: JSZip; xmlString: string }> {
  const { zip, files } = await extractZipFiles(base64, [SLIDE_XML_PATH])
  return { zip, xmlString: files[SLIDE_XML_PATH]! }
}

export async function updateSlideXmlInZip(zip: JSZip, xmlString: string): Promise<string> {
  return await updateZipFiles(zip, { [SLIDE_XML_PATH]: xmlString })
}
