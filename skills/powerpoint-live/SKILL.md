---
name: powerpoint-live
description: "Manipulate live, open PowerPoint presentations on macOS via Office.js MCP bridge. Use when Claude needs to: (1) create, edit, or inspect slides in a running PowerPoint instance, (2) add shapes, text, tables, or formatting to live presentations, (3) capture visual slide screenshots, (4) enable/configure the PowerPoint MCP bridge in a project, (5) execute Office.js code against open presentations. Distinct from the pptx file-editing skill — this works on presentations currently open in PowerPoint."
---

# PowerPoint Live Editing

Edit live, open PowerPoint presentations through an MCP bridge. Changes appear in real-time.

```
Claude Code  ──MCP HTTP (localhost:3001)──>  Bridge Server  ──WS──>  PowerPoint Add-in  ──>  Live Presentation
```

## Setup

When asked to enable or configure PowerPoint MCP in a project — follow the [setup guide](references/setup.md).

## MCP Tools

| Tool | Purpose | Key Parameters |
|------|---------|---------------|
| `list_presentations` | Discover connected presentations | — |
| `get_presentation` | Get slide structure (indices, shape names/types) | `presentationId?` |
| `get_slide` | Get detailed shapes (positions, text, colors) | `slideIndex`, `presentationId?` |
| `get_slide_image` | Capture slide screenshot as PNG | `slideIndex`, `width?` (default 720), `presentationId?` |
| `get_deck_overview` | Visual overview of all/selected slides in one call (thumbnails + text) | `slideRange?`, `imageWidth?` (default 480), `includeImages?`, `presentationId?` |
| `copy_slides` | Copy slides between two open presentations (data stays server-side) | `sourceSlideIndex`, `sourcePresentationId`, `destinationPresentationId`, `formatting?`, `targetSlideId?` |
| `insert_image` | Insert image from file path, URL, or base64 data (data stays server-side for file/url) | `source`, `sourceType` (`file`/`url`/`base64`), `slideIndex?`, `left?`, `top?`, `width?`, `height?`, `presentationId?` |
| `get_local_copy` | Get a local .pptx file path (passthrough for local files, exports cloud files to temp) | `presentationId?` |
| `read_slide_text` | Read raw OOXML `<a:p>` paragraphs from a shape (preserves formatting) | `slideIndex`, `shapeId`, `presentationId?` |
| `edit_slide_text` | Replace paragraph content with raw OOXML (preserves bodyPr/lstStyle) | `slideIndex`, `shapeId`, `xml`, `presentationId?` |
| `read_slide_xml` | Read full slide OOXML or a specific shape's XML | `slideIndex`, `shapeId?`, `presentationId?` |
| `edit_slide_xml` | Replace full slide XML or a specific shape's XML | `slideIndex`, `xml`, `shapeId?`, `presentationId?` |
| `read_slide_zip` | Read multiple files from exported slide zip (slide XML, rels, charts) | `slideIndex`, `paths?`, `presentationId?` |
| `edit_slide_zip` | Update multiple zip files and reimport (auto-registers Content_Types for charts) | `slideIndex`, `files`, `presentationId?` |
| `duplicate_slide` | Clone a slide within the same presentation | `slideIndex`, `insertAfter?`, `presentationId?` |
| `verify_slides` | Check for overlapping, out-of-bounds, empty-text, or tiny shapes | `slideIndex`, `checks?`, `presentationId?` |
| `edit_slide_chart` | Create chart from structured data (generates all OOXML automatically) | `slideIndex`, `chartType`, `title`, `categories`, `series`, `position?`, `options?`, `presentationId?` |
| `execute_officejs` | Run arbitrary Office.js code in the live presentation | `code`, `presentationId?` |

`presentationId` is required only when multiple presentations are connected. Get it from `list_presentations`.

All positioning values from `get_slide` are in **points** (1 pt = 1/72 inch). Standard 16:9 slide: 960 x 540 pt.

## Workflow

1. **Discover**: `list_presentations` — find connected presentations
2. **Audit**: Check existing state — slide count, available layouts, which slides already have content. Use `get_deck_overview` for a visual overview, or `get_presentation` then `get_slide` per slide. This is essential for resuming partial builds or modifying existing decks.
3. **See**: `get_slide_image` — visually inspect specific slides
4. **Modify**: `execute_officejs` — build entire slides in a single call (all shapes, text, connectors, accents at once) for efficiency and to avoid mid-build visual flashing
5. **Verify**: full verification loop (see below)

Always inspect before modifying. Always verify after modifying.

### Verification Loop

After completing work on a slide:

1. **Auto-size first**: set `autoSizeSetting = "AutoSizeShapeToFitText"` on edited text shapes via `execute_officejs` — otherwise `verify_slides` sees stale dimensions
2. **Structural check**: `verify_slides` — overlap, bounds, empty text, tiny shapes
3. **Visual check**: spawn a subagent for independent visual review (see below)
4. **Fix issues** and re-verify until clean

If overlaps/overflow: shorten text, reduce font, reposition body content (not title), or split across slides.

**Intentional overlaps**: When using card patterns (TextBoxes + icons inside RoundedRectangles), `verify_slides` will report many overlaps — these are expected by design. Also, decorative HR lines spanning the full width will overlap with adjacent elements. Only act on overlaps between shapes that should NOT be layered, or on overflow (shapes going off-slide).

**Efficient verification**: For large decks, visually verify only the most complex slides (high shape count, dense content) rather than every slide. Run `verify_slides` on all slides structurally, but pick 4-5 key slides for the visual subagent check.

### Visual Review via Subagent

Use the Agent tool to spawn a subagent that reviews the slide screenshot. The subagent has no conversation context, providing an independent review.

Subagent prompt (replace N with the slide index):
> Call get_slide_image(slideIndex: N) to capture the slide, then review it for: text overflow or truncation, overlapping shapes or text, unreadable text (too small, poor contrast), misalignment or inconsistent spacing, empty or unused space, inconsistent styling (mixed fonts, colors, sizes). Return a JSON array of issues found, each with: severity (error/warning/info), category, description, and suggestion. If no issues found, return [].

Rules: never mention "the reviewer" to user. Speak in first person: "I noticed the title overlaps" not "The reviewer found an overlap." Only use for completed work, not initial inspection.

For `execute_officejs` code patterns, see [code-patterns.md](references/code-patterns.md).

## OOXML Editing Workflow

**Prerequisite**: Load the `/pptx` skill for OOXML structure knowledge (namespaces, element anatomy, formatting rules).

For fine-grained formatting control beyond what Office.js properties expose, use the OOXML tools to read/modify raw slide XML. See [ooxml-reference.md](references/ooxml-reference.md) for detailed tool workflows, batching strategies, unit conversion, and pipeline gotchas.

1. **Discover**: `get_slide(slideIndex)` → find shape IDs
2. **Read**: `read_slide_text` or `read_slide_xml` — get current XML
3. **Modify**: Edit the XML (use `/pptx` skill knowledge)
4. **Write**: `edit_slide_text` or `edit_slide_xml` — apply changes
5. **Verify**: `get_slide_image` — confirm visual result

- `read_slide_text` / `edit_slide_text` — shape-level paragraph editing (preserves `<a:bodyPr>` and `<a:lstStyle>`)
- `read_slide_xml` / `edit_slide_xml` — full slide or shape-level XML editing (full control)
- For batch edits (2+ shapes), use full-slide `read/edit_slide_xml` to avoid multiple reimports

## Hard Limitations

Cannot do via Office.js — do not attempt:

- Insert images with precise shape-level control (use `insert_image` tool — positions via Common API, not shape API)
- Add animations or transitions
- Apply gradients, shadows, or effects (solid fills only)

For charts, use `edit_slide_chart` (declarative) or `edit_slide_zip` (raw OOXML). For slide masters/themes, use `execute_officejs` for backgrounds or `get_local_copy` + python-pptx for full theme editing.

## Content & Design Rules

- Font minimum **14pt** everywhere, preferred body **16pt**
- Always explicitly set `font.size` — do not rely on defaults
- Max 3-4 key points per slide with short supporting text
- Prefer more slides with less content over fewer dense slides
- Use full slide area — stretch content to fill, don't leave large margins
- Never use emoji or Unicode symbols as icons — use geometric shapes as icon substitutes

## Slide Layout Recipes

Common visual patterns for building slides. Adapt colors and content to the user's design system.

### Card Grid
RoundedRectangle as background → TextBox for title (offset ~85pt from left edge for icon space) → TextBox for body below title → Icon (36-48pt) at top-left corner of card.

Calculate card width: `(contentWidth - gaps) / numColumns`. Common configurations: 2x2, 3-across, 4-across, 5-across.

**Intentional overlaps**: Card patterns always report overlaps in `verify_slides` because TextBoxes and icons sit inside the RoundedRectangle. These are expected — only worry about overflow (shapes going off-slide) or unintended sibling overlaps.

### Icon + Text Blocks
Icon (36-48pt) left-aligned → Title TextBox at icon's right → Description TextBox below title, all inside a large RoundedRectangle container. Good for feature lists, "about us" sections, service descriptions.

### Key Numbers / Stats Panel
Large font number (accent color, 28-36pt) + small label below (14-16pt), stacked vertically with separator lines between entries. Good for KPIs, proof points, metrics panels.

### Pillar / Category Map
Vertical tall cards (equal width, evenly spaced) + horizontal bar spanning all pillars at bottom + dashed arrow connectors from each pillar down to the bar. Shows hierarchy: categories above → shared foundation below.

### Left-Right Content Split
Content panel (left, ~45% width) + stats/data panel (right, ~45% width) with a gap between. Good for combining narrative text with data points or proof points.

### Layered Stack
Horizontal rectangles stacked vertically with graduated fill color (darkest at top or bottom). Each layer has a title and description. Shows hierarchy, maturity levels, or technology stacks.

### Before/After Split
Two contrasting colored panels side by side (e.g., muted red for "without" vs green for "with"). Each panel lists bullet points. Optional full-width CTA bar below.

### Case Study / Reference Cards
3 equal-width tall cards, each with: header area (company/project name), description body, and metrics/outcomes section at the bottom.

### Cards with Tier/Tag Badges
Standard cards with a small colored RoundedRectangle "badge" overlaid (e.g., showing a tier level, category label, or status tag). Badge is typically 80-120pt wide, 20-28pt tall, positioned at top-right of the card.

## Gotchas

**XML:**
- Always escape `&` as `&amp;` in `<a:t>` — #1 cause of missing text
- OOXML is fully explicit — every omitted attribute is lost. Copy verbatim from `read_slide_text`.
- No `<!-- -->` comments in code strings — sandbox rejects with `SES_HTML_COMMENT_REJECTED`

**Office.js:**
- Use `getTextFrameOrNullObject()` — never `.textFrame` directly (tables/images/charts throw)
- Loaded values are snapshots — don't branch on stale reads after writes (`hasText` stays stale after setting `textRange.text`)
- No `paragraphs` collection in PowerPoint Office.js
- `slides.add()` always appends — use `slide.moveTo(index)` to reposition
- Always use last master: `masters.items[masters.items.length - 1]` — earlier may be stale
- No `#` prefix for background colors: `{ color: "1A1A1E" }` not `"#1A1A1E"`
- Don't delete placeholders after writing text — `hasText` is stale, you'll delete what you just wrote
- Shape IDs are stable and locale-independent. Shape names change with Office UI language. Always use ID.

**Charts:**
- Always register in `[Content_Types].xml`, include `<c:style val="2"/>`, don't hardcode series colors
- Stacked bars need `<c:overlap val="100"/>`, category axis `majorTickMark val="none"`

**Tables:**
- Height is auto-calculated — `shape.height` and OOXML `<a:ext cy>` are overridden
- Fix overflow via table API only: `cell.font.size` + `row.height`

## Working with python-pptx

For features Office.js cannot access (comments, chart data, embedded objects, master slides, custom XML parts), use `get_local_copy` to get a .pptx file path, then use python-pptx to process it.

- `get_local_copy` returns the existing file path for local files, or exports cloud files to a temp .pptx
- Reads the **saved** state — unsaved changes won't appear until the user saves
- Cached by revision number — only re-exports when the presentation has been saved since last export

## Error Handling

- **"No presentations connected"** — open PowerPoint with the add-in loaded
- **"Multiple presentations connected"** — specify `presentationId`
- **"Add-in disconnected"** — auto-reconnects; wait and retry
- **"Command timed out"** — simplify code or check PowerPoint responsiveness
- **Screenshot via execute_officejs overflows tokens** — always use `get_slide_image` instead (returns MCP image block, not text)
