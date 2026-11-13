/**
 * Classify any Excel chart type — including charts this add-in did not create — by where its
 * series colour lives. Ported from the type lists in `GetStyleColorMode` / `IsLineTarget` /
 * `IsPieChartType` (`modEngineStyle.bas`, `modColorFill.bas`, `modEngineToggles.bas`), refined so
 * marker-only and line-with-marker types colour their markers too.
 *
 * Input is the Office.js `Excel.ChartType` string value (e.g. `"LineMarkers"`). Unknown types fall
 * back to `fill`, as the VBA did.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { ColourTarget } from "../config/chartSpecs";

const POINT_TYPES: readonly string[] = [
  "Pie",
  "PieExploded",
  "3DPie",
  "3DPieExploded",
  "PieOfPie",
  "BarOfPie",
  "Doughnut",
  "DoughnutExploded",
  "Treemap",
  "Sunburst",
];

const LINE_TYPES: readonly string[] = [
  "Line",
  "LineStacked",
  "LineStacked100",
  "3DLine",
  "XYScatterLinesNoMarkers",
  "XYScatterSmoothNoMarkers",
  "Radar",
];

const LINE_MARKER_TYPES: readonly string[] = [
  "LineMarkers",
  "LineMarkersStacked",
  "LineMarkersStacked100",
  "XYScatterLines",
  "XYScatterSmooth",
  "RadarMarkers",
];

const MARKER_TYPES: readonly string[] = ["XYScatter"];

/** Where the series colour of a chart of this type lives. */
export function colourTargetForChartType(chartType: string): ColourTarget {
  if (POINT_TYPES.includes(chartType)) return "point";
  if (LINE_TYPES.includes(chartType)) return "line";
  if (LINE_MARKER_TYPES.includes(chartType)) return "lineMarker";
  if (MARKER_TYPES.includes(chartType)) return "marker";
  return "fill";
}

/** True for pie and doughnut types (the round charts with a square plot and per-slice colours). */
export function isPieChartType(chartType: string): boolean {
  return (
    chartType === "Pie" ||
    chartType === "PieExploded" ||
    chartType === "Doughnut" ||
    chartType === "DoughnutExploded"
  );
}

/**
 * How Apply Chart Style treats a chart type it has no builder spec for — the buckets of the VBA
 * `ClassifyChart` (`modEngineStyle.bas`):
 *   - `pie`         round charts: no axes, legend, per-slice colours
 *   - `axisFill`    2-D bar/column/area/bubble: axes, gridlines, per-series fills
 *   - `lineScatter` 2-D line and scatter: axes, gridlines, colour on lines/markers
 *   - `other`       everything else (3-D, stock, radar, surface, chartex, unknown): chrome and
 *                   best-effort colour only; axes and legend are left as they are
 */
export type StyleClass = "pie" | "axisFill" | "lineScatter" | "other";

const AXIS_FILL_TYPES: readonly string[] = [
  "ColumnClustered",
  "ColumnStacked",
  "ColumnStacked100",
  "BarClustered",
  "BarStacked",
  "BarStacked100",
  "Area",
  "AreaStacked",
  "AreaStacked100",
  "Bubble",
];

export function styleClassForChartType(chartType: string): StyleClass {
  if (isPieChartType(chartType)) return "pie";
  if (AXIS_FILL_TYPES.includes(chartType)) return "axisFill";
  if (/^(Line|XYScatter)/.test(chartType)) return "lineScatter";
  return "other";
}

/** True for the XY scatter family (two value axes), as opposed to line charts. */
export function isScatterChartType(chartType: string): boolean {
  return chartType.startsWith("XYScatter");
}
