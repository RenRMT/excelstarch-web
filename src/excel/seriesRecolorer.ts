/* global Excel */
/**
 * Colour writers for chart series and points. A chart's colour does not always live on the series
 * fill: line charts carry it on the line, scatter charts on the markers, and pie/treemap charts on
 * each point. `ColourTarget` (from `config/chartSpecs`, classified by `logic/chartClass`) says
 * which; these functions write to the right place. Ported from the FILL/LINE modes of
 * `FormatSeriesColors` and the marker colouring in `modChartScatter.bas`.
 *
 * QUEUE host writes only — never a `context.sync()` here. Callers load counts in a prior sync.
 *
 * Office.js `ChartFill` has no transparency control (unlike the VBA `.Format.Fill`), so fills are
 * solid only.
 *
 * INTEROP: Office.js chart API.
 */
import type { Hex } from "../config/brand";
import type { ColourTarget } from "../config/chartSpecs";

/** Colour targets that apply to a whole series (everything except per-point colouring). */
export type SeriesColourTarget = Exclude<ColourTarget, "point">;

/**
 * Set or remove the colour of one series. `hex === null` removes it (the NONE/NOFILL/OFF payload):
 * the fill is cleared, the line hidden, or the markers hidden, depending on the target.
 * `lineWeight` is applied to line targets when given (chart creation uses it; recolouring does not
 * touch the weight).
 */
export function setSeriesColour(
  series: Excel.ChartSeries,
  target: SeriesColourTarget,
  hex: Hex | null,
  lineWeight?: number
): void {
  if (target === "fill") {
    if (hex === null) {
      series.format.fill.clear();
    } else {
      series.format.fill.setSolidColor(hex);
    }
    return;
  }

  if (target === "line" || target === "lineMarker") {
    const line = series.format.line;
    if (hex === null) {
      line.lineStyle = "None";
    } else {
      line.lineStyle = "Continuous";
      line.color = hex;
      if (lineWeight !== undefined) {
        line.weight = lineWeight;
      }
    }
  }

  if (target === "marker" || target === "lineMarker") {
    if (hex === null) {
      series.markerStyle = "None";
    } else {
      series.markerBackgroundColor = hex;
      series.markerForegroundColor = hex;
    }
  }
}

/** Set or remove the fill of one data point (pie slice, treemap tile, single bar). */
export function setPointColour(point: Excel.ChartPoint, hex: Hex | null): void {
  if (hex === null) {
    point.format.fill.clear();
  } else {
    point.format.fill.setSolidColor(hex);
  }
}

/**
 * Apply a precomputed list of colours, index-aligned: to series 1..n, or — for a `point` target —
 * to points 1..n of the first series.
 */
export function applyColours(
  chart: Excel.Chart,
  target: ColourTarget,
  colours: readonly (Hex | null)[],
  lineWeight?: number
): void {
  if (target === "point") {
    const points = chart.series.getItemAt(0).points;
    colours.forEach((hex, i) => setPointColour(points.getItemAt(i), hex));
    return;
  }
  colours.forEach((hex, i) =>
    setSeriesColour(chart.series.getItemAt(i), target, hex, lineWeight)
  );
}

/** Apply one colour (or removal) to a single element: series `index`, or point `index` of series 1. */
export function applyOneColour(
  chart: Excel.Chart,
  target: ColourTarget,
  index: number,
  hex: Hex | null
): void {
  if (target === "point") {
    setPointColour(chart.series.getItemAt(0).points.getItemAt(index), hex);
  } else {
    setSeriesColour(chart.series.getItemAt(index), target, hex);
  }
}
