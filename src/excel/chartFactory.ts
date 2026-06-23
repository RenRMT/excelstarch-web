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

/** Map the UI's ChartKind to the concrete Office.js chart type. */
function chartTypeFor(kind: ChartKind): Excel.ChartType {
  return kind === "bar" ? Excel.ChartType.barClustered : Excel.ChartType.columnClustered;
}

/** Create a new chart from the selected range. */
export function createChart(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  kind: ChartKind
): Excel.Chart {
  return sheet.charts.add(
    chartTypeFor(kind),
    ctx.workbook.getSelectedRange(),
    Excel.ChartSeriesBy.columns
  );
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
