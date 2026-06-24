/* global Excel */
/**
 * Resolve-or-create the target chart and inset it into the plot band — ported from
 * `GetTargetChart` + `PositionChartExChart`. If a chart is currently selected, retype it in place
 * (Office.js `chart.chartType` is settable) so a re-run restyles the same chart and the chrome
 * snaps back to its canvas origin; otherwise create a new chart from the selected range.
 *
 * INTEROP: Office.js chart API.
 */
import type { ChartKind } from "../logic/chartType";
import { chartBandPosition } from "../chrome/chromeLayout";

/**
 * Map the UI's ChartKind to the concrete Office.js chart type. Exhaustive `switch` — adding a new
 * ChartKind without a case here is a compile error (the `never` default fails to type-check).
 * `scatter` uses `xyscatter` (markers only), the closest classic match to the VBA scatter.
 */
function chartTypeFor(kind: ChartKind): Excel.ChartType {
  switch (kind) {
    case "bar":
      return Excel.ChartType.barClustered;
    case "column":
      return Excel.ChartType.columnClustered;
    case "line":
      return Excel.ChartType.line;
    case "area":
      return Excel.ChartType.area;
    case "scatter":
      return Excel.ChartType.xyscatter;
    case "pie":
      return Excel.ChartType.pie;
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
}

/**
 * Resolve the range to chart from the user's selection: a single selected cell is expanded to its
 * surrounding region (the contiguous block / Excel Table around it — matching Excel's own
 * "chart from one cell" behaviour); a multi-cell selection is used as-is. The caller must have
 * loaded `selected.cellCount` (synced) before calling.
 */
export function resolveChartRange(selected: Excel.Range): Excel.Range {
  return selected.cellCount <= 1 ? selected.getSurroundingRegion() : selected;
}

/** Create a new chart from the given (already-resolved) range. */
export function createChart(
  sheet: Excel.Worksheet,
  range: Excel.Range,
  kind: ChartKind
): Excel.Chart {
  return sheet.charts.add(chartTypeFor(kind), range, Excel.ChartSeriesBy.columns);
}

/** Retype an existing chart in place (restyle path). */
export function retypeChart(chart: Excel.Chart, kind: ChartKind): void {
  chart.chartType = chartTypeFor(kind);
}

/**
 * Position the CHART OBJECT as a band inset into the canvas (we size the chart, not an internal
 * plot area — Office.js `Excel.Chart` has none). Bar/column show both axes and no legend.
 */
export function positionChartIntoBand(
  chart: Excel.Chart,
  baseLeft: number,
  baseTop: number,
  showY: boolean,
  showX: boolean,
  hasLegend: boolean
): void {
  const band = chartBandPosition(baseLeft, baseTop, showY, showX, hasLegend);
  chart.left = band.left;
  chart.top = band.top;
  chart.width = band.width;
  chart.height = band.height;
}
