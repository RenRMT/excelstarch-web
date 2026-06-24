/* global Excel */
/**
 * Brand-palette series fill — the FILL-mode apply loop from `FormatSeriesColors`
 * (`modColorSeries.bas`), consuming the already-ported pure `getPaletteColor`. Series 1..count get
 * the brand data colours in palette order; >8 falls back to the neutral (handled in getPaletteColor).
 *
 * QUEUES one `setSolidColor` per series/point — never a `context.sync()` inside the loop. The caller
 * must have loaded the relevant count in a prior sync and passes it in.
 *
 * INTEROP: Office.js chart API.
 */
import { getPaletteColor } from "../logic/colorSeries";

export function colorSeriesByPalette(chart: Excel.Chart, seriesCount: number): void {
  for (let i = 1; i <= seriesCount; i++) {
    // getItemAt is 0-based; getPaletteColor is 1-based (series index).
    chart.series.getItemAt(i - 1).format.fill.setSolidColor(getPaletteColor(i));
  }
}

/**
 * Brand-palette POINT fill for single-series-per-category types (pie/doughnut): a pie has one series
 * whose slices are data POINTS, so colouring the series solid would flatten the whole pie to one
 * colour. Colour each point in palette order instead. The caller must have loaded the first series'
 * `points` "count" in a prior sync and passes it in.
 */
export function colorPointsByPalette(chart: Excel.Chart, pointCount: number): void {
  const points = chart.series.getItemAt(0).points;
  for (let i = 1; i <= pointCount; i++) {
    // getItemAt is 0-based; getPaletteColor is 1-based (point index).
    points.getItemAt(i - 1).format.fill.setSolidColor(getPaletteColor(i));
  }
}
