/* global Excel */
/**
 * Brand-palette series fill — the FILL-mode apply loop from `FormatSeriesColors`
 * (`modColorSeries.bas`), consuming the already-ported pure `getPaletteColor`. Series 1..count get
 * the brand data colours in palette order; >8 falls back to the neutral (handled in getPaletteColor).
 *
 * QUEUES one `setSolidColor` per series — never a `context.sync()` inside the loop. The caller must
 * have loaded `chart.series` "count" in a prior sync and passes it in.
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
