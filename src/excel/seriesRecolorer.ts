/* global Excel */
/**
 * Series fill apply loops — generalizes `colorSeriesByPalette` to apply any precomputed list of
 * fills (the FILL-mode apply from `FormatSeriesColors`), plus a single-series apply for the
 * per-element fill (replacing the VBA Selection model). The ordering math is pure
 * (`logic/seriesFills`); these functions only write to the host.
 *
 * QUEUE host writes only — never a `context.sync()` inside the loop. The caller loads
 * `chart.series` "count" in a prior sync and passes the precomputed fills.
 *
 * Fill-only: bar/column colour the series FILL. When line/scatter land (chart-types breadth PR),
 * an `isLineTarget(kind)` branch writes `series.format.line.color` instead of `fill` — that seam is
 * marked below. Office.js `ChartFill` has no transparency control (unlike the VBA `.Format.Fill`),
 * so series fills are solid only; transparency lives on worksheet *shape* fills, not chart series.
 *
 * INTEROP: Office.js chart API.
 */
import type { Hex } from "../config/brand";

/** Apply a precomputed list of fills to series 1..fills.length (index-aligned, 0-based getItemAt). */
export function applySeriesFills(chart: Excel.Chart, fills: readonly Hex[]): void {
  fills.forEach((hex, i) => {
    // SEAM: line/scatter (future) would set series.getItemAt(i).format.line.color here instead.
    chart.series.getItemAt(i).format.fill.setSolidColor(hex);
  });
}

/**
 * Set or clear the fill of one series (0-based index). `hex === null` removes the fill
 * (the NONE/NOFILL/OFF payload); otherwise it sets the solid colour.
 */
export function applyOneSeriesFill(chart: Excel.Chart, index: number, hex: Hex | null): void {
  const fill = chart.series.getItemAt(index).format.fill;
  if (hex === null) {
    fill.clear();
  } else {
    fill.setSolidColor(hex);
  }
}

/** Clear/set the fill across all series to a single colour (or remove it when `hex === null`). */
export function applyAllSeriesFill(chart: Excel.Chart, seriesCount: number, hex: Hex | null): void {
  for (let i = 0; i < seriesCount; i++) {
    applyOneSeriesFill(chart, i, hex);
  }
}
