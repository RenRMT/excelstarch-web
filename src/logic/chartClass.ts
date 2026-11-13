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
