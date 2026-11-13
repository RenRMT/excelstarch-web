/* global Excel */
/**
 * Resolve-or-create the target chart and inset it into the plot band — ported from
 * `GetTargetChart` + `PositionChartExChart`. If a chart is currently selected, retype it in place
 * (Office.js `chart.chartType` is settable) so a re-run restyles the same chart and the chrome
 * snaps back to its canvas origin; otherwise create a new chart from the selected range.
 *
 * The ChartKind → Excel chart type mapping lives in `config/chartSpecs`.
 *
 * INTEROP: Office.js chart API.
 */
import type { ChartKind } from "../logic/chartType";
import { chartSpecs } from "../config/chartSpecs";
import { chartBandPosition } from "../chrome/chromeLayout";

/** The concrete Office.js chart type for a kind (the spec stores its string value). */
function chartTypeFor(kind: ChartKind): Excel.ChartType {
  return chartSpecs[kind].excelType as Excel.ChartType;
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
 * Position the CHART OBJECT as a band inset into the canvas. The band leaves room above for the
 * title block (and the y-axis title strip when `showY`) and below for the source box (and the
 * x-axis title strip when `showX`). No strip is reserved for the legend: the legend is drawn inside
 * the chart object itself.
 */
export function positionChartIntoBand(
  chart: Excel.Chart,
  baseLeft: number,
  baseTop: number,
  showY: boolean,
  showX: boolean
): void {
  const band = chartBandPosition(baseLeft, baseTop, showY, showX, /*hasLegend*/ false);
  chart.left = band.left;
  chart.top = band.top;
  chart.width = band.width;
  chart.height = band.height;
}
