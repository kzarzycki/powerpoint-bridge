import { describe, expect, it } from 'vitest'
import {
  buildChartRelationship,
  buildChartXml,
  buildGraphicFrame,
  type ChartPosition,
  type ChartSeries,
  pointsToEmu,
  resolveChartPosition,
} from './chart-builder.ts'

describe('buildChartXml', () => {
  const categories = ['Q1', 'Q2', 'Q3']
  const series: ChartSeries[] = [{ name: 'Revenue', values: [100, 150, 120] }]

  it('generates valid column chart XML', () => {
    const xml = buildChartXml('column', 'Sales', categories, series)
    expect(xml).toContain('<?xml version="1.0"')
    expect(xml).toContain('<c:chartSpace')
    expect(xml).toContain('<c:style val="2"/>')
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:barDir val="col"/>')
    expect(xml).toContain('<c:grouping val="clustered"/>')
    expect(xml).toContain('<a:t>Sales</a:t>')
  })

  it('generates bar chart with horizontal direction', () => {
    const xml = buildChartXml('bar', 'Bars', categories, series)
    expect(xml).toContain('<c:barDir val="bar"/>')
  })

  it('generates stacked bar chart with overlap', () => {
    const xml = buildChartXml('column', 'Stacked', categories, series, { stacked: true })
    expect(xml).toContain('<c:grouping val="stacked"/>')
    expect(xml).toContain('<c:overlap val="100"/>')
  })

  it('generates line chart', () => {
    const xml = buildChartXml('line', 'Line', categories, series)
    expect(xml).toContain('<c:lineChart>')
    expect(xml).toContain('<c:grouping val="standard"/>')
  })

  it('generates area chart', () => {
    const xml = buildChartXml('area', 'Area', categories, series)
    expect(xml).toContain('<c:areaChart>')
  })

  it('generates pie chart without axes', () => {
    const xml = buildChartXml('pie', 'Pie', categories, series)
    expect(xml).toContain('<c:pieChart>')
    expect(xml).not.toContain('<c:catAx>')
    expect(xml).not.toContain('<c:valAx>')
  })

  it('generates doughnut chart with hole size', () => {
    const xml = buildChartXml('doughnut', 'Donut', categories, series)
    expect(xml).toContain('<c:doughnutChart>')
    expect(xml).toContain('<c:holeSize val="50"/>')
  })

  it('includes data labels on every series', () => {
    const twoSeries: ChartSeries[] = [
      { name: 'A', values: [10, 20] },
      { name: 'B', values: [30, 40] },
    ]
    const xml = buildChartXml('column', 'Test', ['X', 'Y'], twoSeries)
    // Each series should have its own dLbls
    const matches = xml.match(/<c:dLbls>/g)
    expect(matches?.length).toBe(2)
  })

  it('pie chart shows percent and category name in data labels', () => {
    const xml = buildChartXml('pie', 'Pie', categories, series)
    expect(xml).toContain('<c:showPercent val="1"/>')
    expect(xml).toContain('<c:showCatName val="1"/>')
  })

  it('includes legend by default with font size 1400', () => {
    const xml = buildChartXml('column', 'Test', categories, series)
    expect(xml).toContain('<c:legend>')
    expect(xml).toContain('<c:legendPos val="t"/>')
    expect(xml).toContain('sz="1400"')
  })

  it('omits legend when showLegend is false', () => {
    const xml = buildChartXml('column', 'Test', categories, series, { showLegend: false })
    expect(xml).not.toContain('<c:legend>')
  })

  it('respects legendPosition option', () => {
    const xml = buildChartXml('column', 'Test', categories, series, { legendPosition: 'b' })
    expect(xml).toContain('<c:legendPos val="b"/>')
  })

  it('uses font size 1600 for title', () => {
    const xml = buildChartXml('column', 'Test', categories, series)
    expect(xml).toContain('sz="1600"')
  })

  it('escapes XML special characters in title', () => {
    const xml = buildChartXml('column', 'Sales & Marketing', categories, series)
    expect(xml).toContain('Sales &amp; Marketing')
    expect(xml).not.toContain('Sales & Marketing')
  })

  it('escapes XML special characters in categories', () => {
    const xml = buildChartXml('column', 'Test', ['A&B', 'C<D'], series)
    expect(xml).toContain('A&amp;B')
    expect(xml).toContain('C&lt;D')
  })

  it('escapes XML special characters in series names', () => {
    const s: ChartSeries[] = [{ name: 'Revenue & Profit', values: [100] }]
    const xml = buildChartXml('column', 'Test', ['Q1'], s)
    expect(xml).toContain('Revenue &amp; Profit')
  })

  it('hides data labels when showDataLabels is false', () => {
    const xml = buildChartXml('column', 'Test', categories, series, { showDataLabels: false })
    expect(xml).toContain('<c:showVal val="0"/>')
  })

  it('includes category axis with majorTickMark none', () => {
    const xml = buildChartXml('column', 'Test', categories, series)
    expect(xml).toContain('<c:catAx>')
    expect(xml).toContain('<c:majorTickMark val="none"/>')
  })

  it('includes value axis with majorTickMark out', () => {
    const xml = buildChartXml('column', 'Test', categories, series)
    expect(xml).toContain('<c:valAx>')
    expect(xml).toContain('<c:majorTickMark val="out"/>')
  })

  it('bar chart swaps axis positions', () => {
    const xml = buildChartXml('bar', 'Test', categories, series)
    // Category axis should be on the left for bar charts
    expect(xml).toContain('<c:catAx>')
    expect(xml).toContain('<c:axPos val="l"/>')
    // Value axis should be on the bottom
    expect(xml).toContain('<c:axPos val="b"/>')
  })

  it('uses literal data format (not sheet references)', () => {
    const xml = buildChartXml('column', 'Test', categories, series)
    expect(xml).toContain('<c:strLit>')
    expect(xml).toContain('<c:numLit>')
    expect(xml).not.toContain('Sheet1!')
  })

  it('includes multiple series with correct index and order', () => {
    const multi: ChartSeries[] = [
      { name: 'A', values: [1, 2] },
      { name: 'B', values: [3, 4] },
      { name: 'C', values: [5, 6] },
    ]
    const xml = buildChartXml('column', 'Test', ['X', 'Y'], multi)
    expect(xml).toContain('<c:idx val="0"/>')
    expect(xml).toContain('<c:idx val="1"/>')
    expect(xml).toContain('<c:idx val="2"/>')
    expect(xml).toContain('<c:order val="0"/>')
    expect(xml).toContain('<c:order val="1"/>')
    expect(xml).toContain('<c:order val="2"/>')
  })
})

describe('buildGraphicFrame', () => {
  it('generates valid graphic frame XML', () => {
    const xml = buildGraphicFrame('rId3', { x: 1000, y: 2000, cx: 5000, cy: 3000 }, 'Chart 1', 101)
    expect(xml).toContain('<p:graphicFrame')
    expect(xml).toContain('id="101"')
    expect(xml).toContain('name="Chart 1"')
    expect(xml).toContain('r:id="rId3"')
    expect(xml).toContain('x="1000"')
    expect(xml).toContain('y="2000"')
    expect(xml).toContain('cx="5000"')
    expect(xml).toContain('cy="3000"')
  })

  it('escapes chart name', () => {
    const xml = buildGraphicFrame('rId1', { x: 0, y: 0, cx: 0, cy: 0 }, 'Chart "1"', 1)
    expect(xml).toContain('name="Chart &quot;1&quot;"')
  })
})

describe('buildChartRelationship', () => {
  it('generates valid relationship XML', () => {
    const xml = buildChartRelationship('rId3', '../charts/chart1.xml')
    expect(xml).toContain('Id="rId3"')
    expect(xml).toContain('Target="../charts/chart1.xml"')
    expect(xml).toContain('relationships/chart')
  })
})

describe('pointsToEmu', () => {
  it('converts points to EMU', () => {
    expect(pointsToEmu(72)).toBe(914400) // 1 inch
    expect(pointsToEmu(960)).toBe(12192000) // slide width
    expect(pointsToEmu(540)).toBe(6858000) // slide height
  })
})

describe('resolveChartPosition', () => {
  it('returns defaults when no position given', () => {
    const pos = resolveChartPosition()
    expect(pos.x).toBeGreaterThan(0)
    expect(pos.y).toBeGreaterThan(0)
    expect(pos.cx).toBeGreaterThan(0)
    expect(pos.cy).toBeGreaterThan(0)
  })

  it('overrides specific fields', () => {
    const pos = resolveChartPosition({ left: 100, top: 100 })
    expect(pos.x).toBe(pointsToEmu(100))
    expect(pos.y).toBe(pointsToEmu(100))
    // width/height should still be defaults
    expect(pos.cx).toBeGreaterThan(pointsToEmu(100))
  })

  it('overrides all fields', () => {
    const p: ChartPosition = { left: 50, top: 50, width: 500, height: 300 }
    const pos = resolveChartPosition(p)
    expect(pos.x).toBe(pointsToEmu(50))
    expect(pos.y).toBe(pointsToEmu(50))
    expect(pos.cx).toBe(pointsToEmu(500))
    expect(pos.cy).toBe(pointsToEmu(300))
  })
})
