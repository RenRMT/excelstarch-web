/* global Excel */
/**
 * Brand-palette colouring for a freshly built chart — the apply loop from `FormatSeriesColors`
 * (`modColorSeries.bas`) and `ApplySliceColors` (`modChartPie.bas`), consuming the pure
 * `getPaletteColor`. Elements 1..count get the brand data colours in palette order; beyond the
 * eighth they fall back to the neutral (handled in `getPaletteColor`).
 *
 * `count` is the number of series, or — for a `point` target (pie, treemap) — the number of points
 * in the first series. The caller loads it in a prior sync.
 *
 * QUEUES writes only — never a `context.sync()`.
 *
 * INTEROP: Office.js chart API.
 */
import type { ColourTarget } from "../config/chartSpecs";
import { seriesLineWeight } from "../config/geometry";
import { paletteFills } from "../logic/seriesFills";
import { applyColours } from "./seriesRecolorer";

export function colorByPalette(
  chart: Excel.Chart,
  target: ColourTarget,
  count: number,
  useAltOrder = false
): void {
  applyColours(chart, target, paletteFills(count, useAltOrder), seriesLineWeight);
}
