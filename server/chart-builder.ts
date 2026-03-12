// ---------------------------------------------------------------------------
// Chart OOXML Builder — generates chart XML from structured data
// ---------------------------------------------------------------------------

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

export type ChartType = 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'area'

export interface ChartSeries {
  name: string
  values: number[]
}

export interface ChartOptions {
  stacked?: boolean
  showDataLabels?: boolean
  showLegend?: boolean
  legendPosition?: 't' | 'b' | 'l' | 'r'
}

export interface ChartPosition {
  left?: number // points
  top?: number
  width?: number
  height?: number
}

// ---------------------------------------------------------------------------
// XML escaping
// ---------------------------------------------------------------------------

function esc(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ---------------------------------------------------------------------------
// Points → EMU conversion
// ---------------------------------------------------------------------------

const PT_TO_EMU = 12700

export function pointsToEmu(pt: number): number {
  return Math.round(pt * PT_TO_EMU)
}

// ---------------------------------------------------------------------------
// Default chart position (centered on 16:9 slide)
// ---------------------------------------------------------------------------

const DEFAULT_POSITION = {
  left: 140, // ~2" from left edge
  top: 110, // ~1.5" from top
  width: 680, // ~9.4"
  height: 370, // ~5.1"
}

// ---------------------------------------------------------------------------
// Build chart XML
// ---------------------------------------------------------------------------

export function buildChartXml(
  chartType: ChartType,
  title: string,
  categories: string[],
  series: ChartSeries[],
  options?: ChartOptions,
): string {
  const opts = {
    stacked: false,
    showDataLabels: true,
    showLegend: true,
    legendPosition: 't' as const,
    ...options,
  }

  const titleXml = buildTitle(title)
  const plotArea = buildPlotArea(chartType, categories, series, opts)
  const legendXml = opts.showLegend
    ? `<c:legend><c:legendPos val="${opts.legendPosition}"/><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>`
    : ''

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`,
    '<c:style val="2"/>',
    '<c:chart>',
    titleXml,
    '<c:autoTitleDeleted val="0"/>',
    plotArea,
    legendXml,
    '<c:plotVisOnly val="1"/>',
    '</c:chart>',
    '</c:chartSpace>',
  ].join('')
}

// ---------------------------------------------------------------------------
// Chart title
// ---------------------------------------------------------------------------

function buildTitle(title: string): string {
  return [
    '<c:title>',
    '<c:tx><c:rich>',
    '<a:bodyPr/>',
    '<a:lstStyle/>',
    '<a:p>',
    '<a:pPr><a:defRPr sz="1600" b="1"/></a:pPr>',
    `<a:r><a:rPr lang="en-US" sz="1600" b="1"/><a:t>${esc(title)}</a:t></a:r>`,
    '</a:p>',
    '</c:rich></c:tx>',
    '<c:overlay val="0"/>',
    '</c:title>',
  ].join('')
}

// ---------------------------------------------------------------------------
// Plot area with chart type + axes
// ---------------------------------------------------------------------------

function buildPlotArea(
  chartType: ChartType,
  categories: string[],
  series: ChartSeries[],
  opts: Required<ChartOptions>,
): string {
  const chartElement = buildChartTypeElement(chartType, categories, series, opts)
  const needsAxes = chartType !== 'pie' && chartType !== 'doughnut'

  const axes = needsAxes
    ? [buildCategoryAxis(chartType === 'bar' ? 'l' : 'b'), buildValueAxis(chartType === 'bar' ? 'b' : 'l')].join('')
    : ''

  return `<c:plotArea><c:layout/>${chartElement}${axes}</c:plotArea>`
}

// ---------------------------------------------------------------------------
// Chart type element (barChart, lineChart, pieChart, etc.)
// ---------------------------------------------------------------------------

function buildChartTypeElement(
  chartType: ChartType,
  categories: string[],
  series: ChartSeries[],
  opts: Required<ChartOptions>,
): string {
  const seriesXml = series.map((s, i) => buildSeries(i, s, categories, chartType, opts)).join('')

  switch (chartType) {
    case 'column':
    case 'bar': {
      const dir = chartType === 'bar' ? 'bar' : 'col'
      const grouping = opts.stacked ? 'stacked' : 'clustered'
      const overlap = opts.stacked ? '<c:overlap val="100"/>' : ''
      return `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/>${overlap}${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:barChart>`
    }
    case 'line': {
      const grouping = opts.stacked ? 'stacked' : 'standard'
      return `<c:lineChart><c:grouping val="${grouping}"/>${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:lineChart>`
    }
    case 'area': {
      const grouping = opts.stacked ? 'stacked' : 'standard'
      return `<c:areaChart><c:grouping val="${grouping}"/>${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:areaChart>`
    }
    case 'pie':
      return `<c:pieChart>${seriesXml}</c:pieChart>`
    case 'doughnut':
      return `<c:doughnutChart>${seriesXml}<c:holeSize val="50"/></c:doughnutChart>`
  }
}

// ---------------------------------------------------------------------------
// Individual series
// ---------------------------------------------------------------------------

function buildSeries(
  index: number,
  series: ChartSeries,
  categories: string[],
  chartType: ChartType,
  opts: Required<ChartOptions>,
): string {
  const isPieOrDoughnut = chartType === 'pie' || chartType === 'doughnut'

  const catXml = buildCategoryData(categories)
  const valXml = buildValueData(series.values)
  const dLbls = buildDataLabels(opts.showDataLabels, isPieOrDoughnut)

  return [
    '<c:ser>',
    `<c:idx val="${index}"/>`,
    `<c:order val="${index}"/>`,
    `<c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(series.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`,
    dLbls,
    catXml,
    valXml,
    '</c:ser>',
  ].join('')
}

// ---------------------------------------------------------------------------
// Category / value data (using literal data, not sheet references)
// ---------------------------------------------------------------------------

function buildCategoryData(categories: string[]): string {
  const points = categories.map((cat, i) => `<c:pt idx="${i}"><c:v>${esc(cat)}</c:v></c:pt>`).join('')
  return `<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${points}</c:strLit></c:cat>`
}

function buildValueData(values: number[]): string {
  const points = values.map((val, i) => `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`).join('')
  return `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numLit></c:val>`
}

// ---------------------------------------------------------------------------
// Data labels
// ---------------------------------------------------------------------------

function buildDataLabels(show: boolean, isPieOrDoughnut: boolean): string {
  if (!show) {
    return '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>'
  }
  if (isPieOrDoughnut) {
    return '<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls>'
  }
  return '<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>'
}

// ---------------------------------------------------------------------------
// Axes
// ---------------------------------------------------------------------------

function buildCategoryAxis(position: 'b' | 'l'): string {
  return [
    '<c:catAx>',
    '<c:axId val="1"/>',
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    '<c:delete val="0"/>',
    `<c:axPos val="${position}"/>`,
    '<c:majorTickMark val="none"/>',
    '<c:minorTickMark val="none"/>',
    '<c:tickLblPos val="nextTo"/>',
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>',
    '<c:crossAx val="2"/>',
    '</c:catAx>',
  ].join('')
}

function buildValueAxis(position: 'b' | 'l'): string {
  return [
    '<c:valAx>',
    '<c:axId val="2"/>',
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    '<c:delete val="0"/>',
    `<c:axPos val="${position}"/>`,
    '<c:majorTickMark val="out"/>',
    '<c:minorTickMark val="none"/>',
    '<c:tickLblPos val="nextTo"/>',
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>',
    '<c:crossAx val="1"/>',
    '</c:valAx>',
  ].join('')
}

// ---------------------------------------------------------------------------
// Graphic frame for embedding chart in slide XML
// ---------------------------------------------------------------------------

export function buildGraphicFrame(
  rId: string,
  position: { x: number; y: number; cx: number; cy: number }, // EMU
  chartName: string,
  shapeId: number,
): string {
  return [
    '<p:graphicFrame',
    ` xmlns:p="${P_NS}"`,
    ` xmlns:a="${A_NS}"`,
    ` xmlns:r="${R_NS}">`,
    '<p:nvGraphicFramePr>',
    `<p:cNvPr id="${shapeId}" name="${esc(chartName)}"/>`,
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>',
    '<p:nvPr/>',
    '</p:nvGraphicFramePr>',
    `<p:xfrm><a:off x="${position.x}" y="${position.y}"/><a:ext cx="${position.cx}" cy="${position.cy}"/></p:xfrm>`,
    `<a:graphic><a:graphicData uri="${C_NS}">`,
    `<c:chart xmlns:c="${C_NS}" r:id="${rId}"/>`,
    '</a:graphicData></a:graphic>',
    '</p:graphicFrame>',
  ].join('')
}

// ---------------------------------------------------------------------------
// Relationship XML snippet
// ---------------------------------------------------------------------------

export function buildChartRelationship(rId: string, chartPath: string): string {
  return `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${chartPath}"/>`
}

// ---------------------------------------------------------------------------
// Resolve chart position (points → EMU with defaults)
// ---------------------------------------------------------------------------

export function resolveChartPosition(pos?: ChartPosition): {
  x: number
  y: number
  cx: number
  cy: number
} {
  const p = { ...DEFAULT_POSITION, ...pos }
  return {
    x: pointsToEmu(p.left),
    y: pointsToEmu(p.top),
    cx: pointsToEmu(p.width),
    cy: pointsToEmu(p.height),
  }
}
