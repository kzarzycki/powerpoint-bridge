# OOXML Live Editing Reference

**Prerequisite**: Load the `/pptx` skill first for OOXML structure knowledge (namespaces, `<a:p>/<a:r>/<a:rPr>` anatomy, formatting rules, XML escaping). This reference covers how to apply that knowledge through the bridge's live-editing MCP tools.

## When to Use OOXML vs Office.js

| Approach | Best for |
|----------|----------|
| Office.js (`execute_officejs`) | Simple text, shapes, fills, positioning — anything the Office.js API exposes directly |
| OOXML tools (`read/edit_slide_text`, `read/edit_slide_xml`) | Rich text formatting, precise paragraph control, bulk multi-shape edits |
| File-based (`get_local_copy` + `/pptx` skill) | Charts, master/theme editing, rels, Content_Types — anything beyond slide XML |

## Workflow: Discover → Read → Modify → Write → Verify

1. `get_slide(slideIndex)` → find shape IDs (the `id` field on each shape)
2. `read_slide_text(slideIndex, shapeId)` or `read_slide_xml(slideIndex, shapeId?)`
3. Modify the XML using `/pptx` skill knowledge
4. `edit_slide_text(slideIndex, shapeId, xml)` or `edit_slide_xml(slideIndex, xml, shapeId?)`
5. `get_slide_image(slideIndex)` → visual verification

Always inspect before modifying. Always verify after.

## Shape ID Mapping

- `get_slide` returns shapes with an `id` field (e.g. `"5"`)
- This matches `<p:cNvPr id="5">` in the OOXML
- Always use `get_slide` first to discover IDs — don't guess
- Shape IDs may change after reimport (Office.js assigns new IDs on `insertSlidesFromBase64`)

## read_slide_text / edit_slide_text

Paragraph-level editing for a single shape:

- `read_slide_text` returns the `<a:p>` paragraph elements from a shape
- `edit_slide_text` replaces paragraph content — `<a:bodyPr>` and `<a:lstStyle>` are preserved automatically
- You only work with the paragraph XML (the `<a:p>` elements)

```
// Read current paragraphs
read_slide_text(slideIndex: 0, shapeId: "2")
// Returns: <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r></a:p>

// Write modified paragraphs back
edit_slide_text(slideIndex: 0, shapeId: "2", xml: '<a:p>..modified..</a:p>')
```

## read_slide_xml / edit_slide_xml

Full slide or shape-level XML editing:

- **Without shapeId**: returns/replaces the full slide XML (`<p:sld>...</p:sld>`)
- **With shapeId**: returns/replaces that shape's `<p:sp>` element only

```
// Full slide XML
read_slide_xml(slideIndex: 0)
edit_slide_xml(slideIndex: 0, xml: '<p:sld>...</p:sld>')

// Single shape XML
read_slide_xml(slideIndex: 0, shapeId: "5")
edit_slide_xml(slideIndex: 0, shapeId: "5", xml: '<p:sp>...</p:sp>')
```

Use full-slide mode for batch editing multiple shapes in a single reimport.

## Batching Multiple Edits

Each edit tool call triggers a full export → modify → delete → reimport cycle:

- Multiple `edit_slide_text` calls on the same slide = multiple reimports (visible flashing)
- **For 2+ shapes on the same slide**: use `read_slide_xml` (full slide, no shapeId) → modify all shapes in the XML → `edit_slide_xml` (full slide) — single reimport, no flashing

```
// Bad: 3 reimports, visible flashing
edit_slide_text(slideIndex: 0, shapeId: "2", xml: '...')
edit_slide_text(slideIndex: 0, shapeId: "5", xml: '...')
edit_slide_text(slideIndex: 0, shapeId: "7", xml: '...')

// Good: 1 reimport, no flashing
xml = read_slide_xml(slideIndex: 0)          // full slide
// modify shapes 2, 5, 7 in the XML
edit_slide_xml(slideIndex: 0, xml: modified)  // single reimport
```

## Units: Points vs EMU

| Context | Unit | 1 inch = |
|---------|------|----------|
| `get_slide` / Office.js | Points | 72 pt |
| OOXML (`<a:off>`, `<a:ext>`) | EMU | 914,400 EMU |

Conversion: **EMU = points × 12,700**

| Reference | Points | EMU |
|-----------|--------|-----|
| Standard 16:9 slide width | 960 pt | 12,192,000 |
| Standard 16:9 slide height | 540 pt | 6,858,000 |
| 1 inch | 72 pt | 914,400 |
| 1 cm | 28.35 pt | 360,000 |

When moving positions from `get_slide` output into OOXML, multiply by 12,700.

## Export/Reimport Mechanics

The bridge handles the export/reimport cycle transparently — you just send/receive XML. Under the hood:

1. **Export**: `slide.exportAsBase64()` → single-slide .pptx as Base64
2. **Unzip**: Server extracts `ppt/slides/slide1.xml` (always this path, regardless of slideIndex)
3. **Modify**: Server applies your XML changes to the extracted slide
4. **Repack**: Server creates a new Base64 .pptx with the modified XML
5. **Delete**: Original slide is deleted from the presentation
6. **Reimport**: `presentation.insertSlidesFromBase64()` at the same position

The data stays server-side — XML content never enters Claude's context.

## Multi-File Zip Access: read_slide_zip / edit_slide_zip

`read_slide_xml` / `edit_slide_xml` only access slide XML. For charts, rels, or Content_Types, use the zip-level tools:

```
// Discover all files in the exported zip
read_slide_zip(slideIndex: 0)
// Returns: { zipContents: { path: content, ... }, allPaths: [...] }

// Read specific files
read_slide_zip(slideIndex: 0, paths: ["ppt/charts/chart1.xml", "ppt/slides/_rels/slide1.xml.rels"])

// Update multiple files in one reimport (can add new files)
edit_slide_zip(slideIndex: 0, files: {
  "ppt/slides/slide1.xml": modifiedSlideXml,
  "ppt/charts/chart1.xml": chartXml,
  "ppt/slides/_rels/slide1.xml.rels": updatedRels
})
```

**Auto Content_Types**: When `edit_slide_zip` adds new files under `ppt/charts/`, it auto-registers them in `[Content_Types].xml`. You can still include `[Content_Types].xml` explicitly in the files map to override.

## Current Tool Limitations

The zip-level tools access all text/XML files in the **single-slide export**. They **cannot** access:

- **Masters/themes** — not included in single-slide export (need full pptx)
- **Binary media** — `ppt/media/` files are binary, not text (use `insert_image` instead)
- **Notes** — `ppt/notesSlides/` may not be included in single-slide export

**Workaround**: Use `get_local_copy` to get a .pptx file path, then edit with the `/pptx` skill's file-based workflow.

## Charts via OOXML

Use `read_slide_zip` / `edit_slide_zip` for chart creation and editing. Chart creation requires:

### Chart XML structure (`ppt/charts/chartN.xml`)

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <!-- Chart type element: barChart, lineChart, pieChart, etc. -->
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>
```

### Chart types

| OOXML element | Chart type |
|---------------|------------|
| `<c:barChart>` | Bar/column |
| `<c:lineChart>` | Line |
| `<c:pieChart>` | Pie |
| `<c:areaChart>` | Area |
| `<c:scatterChart>` | Scatter/XY |
| `<c:doughnutChart>` | Doughnut |

### Registration required

1. **Content_Types** — `edit_slide_zip` auto-registers this when adding `ppt/charts/*.xml` files
2. **Slide rels** — add `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>` to `ppt/slides/_rels/slide1.xml.rels`
3. **Graphic frame** on slide — add `<p:graphicFrame>` with `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId3"/></a:graphicData></a:graphic>` to the slide XML

### Chart Rules

- Every chart MUST include: `<c:title>`, `<c:legend>` (with `legendPos val="t"` and `overlay val="0"`), `<c:dLbls>` on every series
- Always include `<c:style val="2"/>` in `<c:chartSpace>` for theme color inheritance
- Do NOT hardcode series colors with `<c:spPr>` — omit so theme accents apply
- Stacked bar/column: MUST add `<c:overlap val="100"/>` inside `<c:barChart>` — without it, bars render side by side
- Category axis: `<c:majorTickMark val="none"/>` (ticks fall between categories and look offset)
- Value axis: `<c:majorTickMark val="out"/>` (keep major ticks visible)
- Both axes: `<c:minorTickMark val="none"/>`
- Chart font minimums: title sz="1600" (16pt), axis labels sz="1400" (14pt), data labels sz="1400" (14pt), legend sz="1400" (14pt)
- Pie/doughnut: add `<c:showPercent val="1"/>` and optionally `<c:showCatName val="1"/>` to `<c:dLbls>`

### Data Labels

Always add to each `<c:ser>`:
```xml
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>
```

### Full Chart Creation Example

```
// Using edit_slide_zip for chart creation (MCP bridge approach):

// 1. Read the slide zip
read_slide_zip(slideIndex: 0)

// 2. Prepare chart XML
var chartXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<c:style val="2"/>' +
  '<c:chart>' +
  '  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Chart Title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '  <c:plotArea>' +
  '    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
  '      <c:ser><c:idx val="0"/><c:tx><c:v>Series</c:v></c:tx>' +
  '        <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>' +
  '        <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>150</c:v></c:pt><c:pt idx="2"><c:v>120</c:v></c:pt></c:numLit></c:val>' +
  '        <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>' +
  '      </c:ser>' +
  '      <c:axId val="1"/><c:axId val="2"/>' +
  '    </c:barChart>' +
  '    <c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/></c:catAx>' +
  '    <c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>' +
  '  </c:plotArea>' +
  '  <c:legend><c:legendPos val="t"/><c:overlay val="0"/></c:legend>' +
  '</c:chart></c:chartSpace>';

// 3. Prepare updated slide XML with graphic frame (positions in EMU)
// Add to <p:spTree>:
// <p:graphicFrame>
//   <p:nvGraphicFramePr><p:cNvPr id="101" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
//   <p:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="5080000" cy="3810000"/></p:xfrm>
//   <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
//     <c:chart r:id="rId3"/>
//   </a:graphicData></a:graphic>
// </p:graphicFrame>

// 4. Prepare updated rels with chart relationship
// Add: <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>

// 5. Write all files in one call (Content_Types auto-registered for charts)
edit_slide_zip(slideIndex: 0, files: {
  "ppt/slides/slide1.xml": modifiedSlideXml,
  "ppt/charts/chart1.xml": chartXml,
  "ppt/slides/_rels/slide1.xml.rels": updatedRelsXml
})
```

## Master/Theme via OOXML (Future Reference)

When full pptx export is added, theme editing involves:

### Theme structure (`ppt/theme/theme1.xml`)

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <!-- accent2-6, hlink, folHlink -->
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
```

### What execute_officejs CAN do now

- Set slide backgrounds: `slide.fill.setSolidColor("#hex")`
- Read some theme colors via shapes
- Apply formatting that references theme colors

### What needs full pptx export

- Full color scheme editing (`<a:clrScheme>`)
- Font scheme changes (`<a:fontScheme>`)
- Master slide layout editing (`ppt/slideMasters/`, `ppt/slideLayouts/`)
- Background styles with gradients or patterns

## Text Formatting Rules

### The Core Rule

**Never use Office.js to read text content.** `textRange.text` returns plain text — all formatting (bold, font size, color, bullets) is stripped. Use `read_slide_text` to read. Office.js is only for shape metadata (IDs, positions, dimensions) and simple plain-text writes.

### DOs

- Always call `read_slide_text` before `edit_slide_text` to see the existing XML
- Copy every `<a:p>` block verbatim from the read output, then make only the specific change needed
- Copy formatting from similar paragraphs when adding new content — new bullets should use the same `<a:pPr>` and `<a:rPr>` as existing ones
- Use `<a:buChar>` in `<a:pPr>` for native PowerPoint bullets
- Keep theme colors (`<a:schemeClr>`) — never replace with hardcoded hex unless explicitly asked
- Escape XML special characters in `<a:t>`: `&` → `&amp;`, `<` → `&lt;`, `>` → `&gt;`

### DON'Ts

- **Don't put bare `&` in `<a:t>`** — always escape as `&amp;`. This is the #1 cause of missing text. `Sales & Marketing` must be `Sales &amp; Marketing`
- Don't rewrite or "clean up" XML — copy verbatim. If read returns `<a:rPr lang="en-US" sz="1000" b="1" dirty="0">`, write exactly that
- Don't use `lvl="1"` for top-level bullets — lvl is 0-based: top-level = `lvl="0"` or omit lvl. `lvl="1"` creates sub-bullets
- When editing, copy existing `<a:pPr>` (which may use explicit marL/indent) rather than inventing new attributes
- Don't put the `•` character in `<a:t>` — use `<a:buChar char="•"/>` in `<a:pPr>`
- Don't mix header and bullet formatting — headers use `<a:buNone/>` with different attributes

## Pipeline-Specific Gotchas

1. **Shape IDs change after reimport** — Office.js assigns new IDs on `insertSlidesFromBase64`. Always re-read `get_slide` after editing if you need to reference shapes again.

2. **Edit slides in reverse index order** — each reimport deletes and reinserts the slide. If editing slides 0, 1, 2, edit in order 2 → 1 → 0 to avoid index shifting.

3. **Namespace variations** — `read_slide_xml` returns the exported slide's XML verbatim. Namespace prefixes may differ slightly from a raw .pptx file due to Office.js export behavior. Match what you read, don't assume canonical prefixes.

4. **Single-slide export scope** — the exported zip always contains just one slide at `ppt/slides/slide1.xml`, even if the original was slide 5 in the deck. Shape references to external content (hyperlinks, charts, media) may break if rels aren't included.

5. **Reimport is destructive** — the original slide is deleted before reimport. If the modified XML is malformed, the slide may be lost. Always keep the read XML as a fallback reference.

6. **Concurrent edits** — if the user edits the slide in PowerPoint while you're modifying XML, the reimport will overwrite their changes. Warn users before batch OOXML operations.
