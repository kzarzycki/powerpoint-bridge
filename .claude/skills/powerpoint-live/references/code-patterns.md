# Office.js Code Patterns

All code passed to `execute_officejs` runs inside `PowerPoint.run(async (context) => { ... })`.
The `context` variable is pre-bound. Always call `await context.sync()` after loading properties or batching changes.

## Slides

```javascript
// Add blank slide
context.presentation.slides.add();
await context.sync();

// Delete slide by index
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
slides.items[2].delete();
await context.sync();

// Get slide count
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
return slides.items.length;
```

## Geometric Shapes

```javascript
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
var shapes = slides.items[0].shapes;

// Add rectangle
var rect = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
rect.left = 100; rect.top = 100; rect.width = 200; rect.height = 150;
rect.fill.setSolidColor("#2196F3");
await context.sync();
```

Available types: `rectangle`, `roundedRectangle`, `ellipse`, `triangle`, `diamond`, `parallelogram`, `trapezoid`, `pentagon`, `hexagon`, `heptagon`, `octagon`, `decagon`, `dodecagon`, `star4`, `star5`, `star6`, `star8`, `star10`, `star12`, `star16`, `star24`, `star32`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `plus`, `heart`, `cloud`, and more.

## Text Boxes & Lines

```javascript
// Text box with position
shapes.addTextBox("Hello World", { left: 50, top: 50, width: 300, height: 40 });
await context.sync();

// Straight line
shapes.addLine(PowerPoint.ConnectorType.straight, {
  left: 50, top: 200, width: 400, height: 0
});
await context.sync();
```

## Text Formatting

```javascript
// Set text and color
shape.textFrame.textRange.text = "Styled text";
shape.textFrame.textRange.font.color = "#FF5722";
shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
await context.sync();

// Multiline
shape.textFrame.textRange.text = "Line 1\nLine 2\nLine 3";
await context.sync();
```

## Fill & Color

Only solid fills supported — no gradients, shadows, or effects.

```javascript
shape.fill.setSolidColor("#4CAF50");
// Named colors work too
shape.fill.setSolidColor("lightblue");
shape.fill.setSolidColor("coral");
await context.sync();
```

## Grouping

```javascript
var group = shapes.addGroup(arrayOfShapes);
await context.sync();

// Ungroup
group.group.ungroup();
await context.sync();
```

## Tables (API 1.8+)

```javascript
// Add 4-row, 3-column table
var table = shapes.addTable(4, 3);
await context.sync();
```

## Screenshots

Always use the `get_slide_image` MCP tool for visual screenshots. Do NOT call `getImageAsBase64` through `execute_officejs` — the raw Base64 text overflows the token limit.

```javascript
// Only if you need the raw Base64 data (prefer get_slide_image tool instead)
var slide = slides.items[0];
var result = slide.getImageAsBase64({ width: 720 });
await context.sync();
return { base64: result.value };

// With custom dimensions (height auto-calculated if omitted)
var result = slide.getImageAsBase64({ width: 1280, height: 720 });
await context.sync();
```

## Copying Slides Between Presentations

Use the `copy_slides` MCP tool to copy slides between two open presentations. The Base64 data transfers server-side (Add-in A → Bridge Server → Add-in B) and never enters Claude's context.

```
copy_slides(
  sourceSlideIndex: 2,
  sourcePresentationId: "deck-a.pptx",
  destinationPresentationId: "deck-b.pptx",
  formatting: "UseDestinationTheme",  // optional
  targetSlideId: "267#"               // optional: insert after this slide
)
```

**Formatting options:**
- `"KeepSourceFormatting"` (default) — inserted slides keep their original theme/colors
- `"UseDestinationTheme"` — inserted slides adopt the target presentation's theme

**Slide ID formats** for `targetSlideId`:
- `"267#"` — slide ID only
- `"#763315295"` — creation ID only
- `"267#763315295"` — both

Under the hood, `copy_slides` calls `slide.exportAsBase64()` on the source and `presentation.insertSlidesFromBase64()` on the destination. For direct use via `execute_officejs`:

```javascript
// Export a slide to Base64 .pptx (API 1.8+)
var slide = slides.items[0];
var result = slide.exportAsBase64();
await context.sync();
return result.value; // Base64 .pptx string

// Insert slides from Base64 .pptx
context.presentation.insertSlidesFromBase64(base64String, {
  formatting: "UseDestinationTheme",
  targetSlideId: "267#"
});
await context.sync();
```

## Reading Content

```javascript
var slide = slides.items[0];
slide.shapes.load("items");
await context.sync();
var texts = [];
for (var i = 0; i < slide.shapes.items.length; i++) {
  var s = slide.shapes.items[i];
  try {
    s.textFrame.load("textRange");
    await context.sync();
    texts.push({ name: s.name, text: s.textFrame.textRange.text });
  } catch (e) { /* no text frame */ }
}
return texts;
```

## Custom Properties (API 1.7+)

```javascript
context.presentation.properties.custom.add("status", "draft");
await context.sync();
```

## Units & Positioning

All values in **points** (1 pt = 1/72 inch).

| Conversion | Formula |
|---|---|
| Inches to points | inches * 72 |
| cm to points | cm / 2.54 * 72 |

Standard 16:9 slide: **960 x 540 pt** (13.33 x 7.5 in)
Standard 4:3 slide: **960 x 720 pt** (13.33 x 10 in)

| Reference | Value |
|---|---|
| Full width | 960 pt |
| Center X | 480 pt |
| Center Y (16:9) | 270 pt |
| Typical margin | 36 pt (0.5 in) |
| Title area | top 36 pt, height ~72 pt |
| Content area (16:9) | top 120 pt to 504 pt |
