import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  autoRegisterContentTypes,
  escapeXml,
  extractParagraphs,
  extractSlideXmlFromZip,
  extractZipFiles,
  findShapeById,
  listZipPaths,
  parseSlideXml,
  replaceParagraphs,
  replaceShape,
  serializeXml,
  updateSlideXmlInZip,
  updateZipFiles,
} from './xml-helpers.ts'

const SAMPLE_SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r></a:p>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>World</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Content 2"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Body text</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`

describe('xml-helpers', () => {
  describe('findShapeById', () => {
    it('finds shape by ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')
      expect(shape).not.toBeNull()
      expect(serializeXml(shape!)).toContain('Title 1')
    })

    it('finds shape with different ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '5')
      expect(shape).not.toBeNull()
      expect(serializeXml(shape!)).toContain('Content 2')
    })

    it('returns null for non-existent ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      expect(findShapeById(doc, '999')).toBeNull()
    })
  })

  describe('extractParagraphs', () => {
    it('extracts <a:p> elements as XML string', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')!
      const result = extractParagraphs(shape)
      expect(result).toContain('<a:p')
      expect(result).toContain('Hello')
      expect(result).toContain('World')
      expect(result).toContain('b="1"')
      // Should NOT contain bodyPr or lstStyle
      expect(result).not.toContain('<a:bodyPr')
      expect(result).not.toContain('<a:lstStyle')
    })

    it('throws for shape without txBody', () => {
      const xml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:nvSpPr><p:cNvPr id="1" name="NoText"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
      </p:sp>`
      const doc = parseSlideXml(xml)
      const shapes = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sp')
      expect(() => extractParagraphs(shapes[0]!)).toThrow('no text body')
    })
  })

  describe('replaceParagraphs', () => {
    it('replaces paragraphs while preserving bodyPr and lstStyle', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')!
      const newXml =
        '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Replaced</a:t></a:r></a:p>'
      replaceParagraphs(doc, shape, newXml)
      const result = serializeXml(doc)
      expect(result).toContain('Replaced')
      expect(result).not.toContain('Hello')
      expect(result).not.toContain('World')
      // bodyPr should still be present
      expect(result).toContain('anchor="ctr"')
    })
  })

  describe('replaceShape', () => {
    it('replaces entire shape element', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '5')!
      const newShapeXml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="5" name="Replaced Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>New content</a:t></a:r></a:p></p:txBody>
      </p:sp>`
      replaceShape(doc, shape, newShapeXml)
      const result = serializeXml(doc)
      expect(result).toContain('Replaced Shape')
      expect(result).toContain('New content')
      expect(result).not.toContain('Body text')
    })
  })

  describe('zip helpers', () => {
    it('extracts and updates slide XML in zip', async () => {
      // Create a test zip with slide XML
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      const base64 = await zip.generateAsync({ type: 'base64' })

      // Extract
      const { zip: extractedZip, xmlString } = await extractSlideXmlFromZip(base64)
      expect(xmlString).toContain('Hello')
      expect(xmlString).toContain('World')

      // Modify and update
      const modified = xmlString.replace('Hello', 'Modified')
      const updatedBase64 = await updateSlideXmlInZip(extractedZip, modified)

      // Verify round-trip
      const { xmlString: finalXml } = await extractSlideXmlFromZip(updatedBase64)
      expect(finalXml).toContain('Modified')
      expect(finalXml).not.toContain('Hello')
    })

    it('throws when slide XML not found in zip', async () => {
      const zip = new JSZip()
      zip.file('ppt/other.xml', '<root/>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      await expect(extractSlideXmlFromZip(base64)).rejects.toThrow('File not found in zip')
    })
  })

  describe('multi-file zip helpers', () => {
    it('listZipPaths returns sorted paths', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', '<slide/>')
      zip.file('[Content_Types].xml', '<Types/>')
      zip.file('ppt/charts/chart1.xml', '<chart/>')
      const paths = listZipPaths(zip)
      expect(paths).toEqual(['[Content_Types].xml', 'ppt/charts/chart1.xml', 'ppt/slides/slide1.xml'])
    })

    it('extractZipFiles extracts specific paths', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('ppt/charts/chart1.xml', '<c:chartSpace/>')
      zip.file('ppt/media/image1.png', 'binarydata')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { files } = await extractZipFiles(base64, ['ppt/slides/slide1.xml', 'ppt/charts/chart1.xml'])
      expect(Object.keys(files)).toHaveLength(2)
      expect(files['ppt/slides/slide1.xml']).toContain('Hello')
      expect(files['ppt/charts/chart1.xml']).toContain('chartSpace')
    })

    it('extractZipFiles throws for missing path', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      const base64 = await zip.generateAsync({ type: 'base64' })

      await expect(extractZipFiles(base64, ['ppt/slides/slide1.xml', 'ppt/missing.xml'])).rejects.toThrow(
        'File not found in zip: ppt/missing.xml',
      )
    })

    it('extractZipFiles auto-discovers text/xml files when no paths given', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types/>')
      zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
      zip.file('ppt/media/image1.png', 'binarydata')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { files } = await extractZipFiles(base64)
      expect(files['ppt/slides/slide1.xml']).toBeDefined()
      expect(files['[Content_Types].xml']).toBeDefined()
      expect(files['ppt/slides/_rels/slide1.xml.rels']).toBeDefined()
      // Binary files should NOT be included
      expect(files['ppt/media/image1.png']).toBeUndefined()
    })

    it('updateZipFiles updates multiple files and adds new ones', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types/>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { zip: extractedZip } = await extractZipFiles(base64, ['ppt/slides/slide1.xml'])
      const updatedBase64 = await updateZipFiles(extractedZip, {
        'ppt/slides/slide1.xml': SAMPLE_SLIDE_XML.replace('Hello', 'Updated'),
        'ppt/charts/chart1.xml': '<c:chartSpace><c:chart/></c:chartSpace>',
      })

      // Verify round-trip
      const { files } = await extractZipFiles(updatedBase64, [
        'ppt/slides/slide1.xml',
        'ppt/charts/chart1.xml',
        '[Content_Types].xml',
      ])
      expect(files['ppt/slides/slide1.xml']).toContain('Updated')
      expect(files['ppt/charts/chart1.xml']).toContain('<c:chart/>')
      expect(files['[Content_Types].xml']).toBe('<Types/>') // untouched
    })
  })

  describe('autoRegisterContentTypes', () => {
    it('adds chart Override when missing', async () => {
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('PartName="/ppt/charts/chart1.xml"')
      expect(ct).toContain('drawingml.chart+xml')
    })

    it('does not duplicate existing Override', async () => {
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      )
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      // Should appear exactly once
      const matches = ct.match(/PartName="\/ppt\/charts\/chart1.xml"/g)
      expect(matches).toHaveLength(1)
    })

    it('ignores paths that do not match known patterns', async () => {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<Types></Types>')
      await autoRegisterContentTypes(zip, ['ppt/slides/slide1.xml', 'ppt/unknown/foo.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toBe('<Types></Types>') // unchanged
    })

    it('handles multiple chart files', async () => {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<Types></Types>')
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('PartName="/ppt/charts/chart1.xml"')
      expect(ct).toContain('PartName="/ppt/charts/chart2.xml"')
    })
  })

  describe('escapeXml', () => {
    it('escapes all XML special characters', () => {
      expect(escapeXml('&')).toBe('&amp;')
      expect(escapeXml('<')).toBe('&lt;')
      expect(escapeXml('>')).toBe('&gt;')
      expect(escapeXml('"')).toBe('&quot;')
      expect(escapeXml("'")).toBe('&apos;')
    })

    it('escapes mixed content', () => {
      expect(escapeXml('Q1 Revenue & "Margins" <2026>')).toBe('Q1 Revenue &amp; &quot;Margins&quot; &lt;2026&gt;')
    })

    it('passes through normal text unchanged', () => {
      expect(escapeXml('Hello World 123')).toBe('Hello World 123')
    })
  })
})
