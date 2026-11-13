/* global Excel */
/**
 * Per-type finishing — the touches the VBA chart builders applied after the shared pipeline: gap
 * width and overlap, tick marks, start-on-axis, shadows, scatter markers, doughnut hole, the square
 * pie plot and data-label fonts. What applies to which kind is described in `config/chartSpecs`.
 *
 * Split in two because Office.js fails a whole batch on one rejected write, and some of these are
 * not accepted by every chart type or host:
 *   - `queueSeriesAndAxisFinishing` — well-supported properties on 2-D classic charts;
 *   - `queueExtraFinishing` — plot-area geometry, doughnut hole and data-label fonts, which chartex
 *     types and some hosts may reject.
 * The caller syncs each in its own guarded sync, so a rejection costs only that group.
 *
 * QUEUES writes only — never a `context.sync()`.
 *
 * INTEROP: Office.js chart API.
 */
import type { Finishing } from "../config/chartSpecs";
import type { ChartDefaults } from "../config/chartDefaults";
import { Axis } from "../config/enums";
import { squarePlotArea, legendHeight } from "../config/geometry";
import { fontPrimary, axisFontSize, axisFontColor } from "../config/fonts";

/** Gap/overlap, shadows and markers on every series; tick marks and start-on-axis on the axes. */
export function queueSeriesAndAxisFinishing(
  chart: Excel.Chart,
  finishing: Finishing,
  defaults: ChartDefaults,
  seriesCount: number
): void {
  for (let i = 0; i < seriesCount; i++) {
    const series = chart.series.getItemAt(i);
    if (finishing.gapWidth !== undefined) series.gapWidth = finishing.gapWidth;
    if (finishing.overlap !== undefined) series.overlap = finishing.overlap;
    if (finishing.noShadow) series.showShadow = false;
    if (finishing.markers) {
      const styles = finishing.markers.styles;
      series.markerStyle = styles[i % styles.length] as Excel.ChartMarkerStyle;
      series.markerSize = finishing.markers.size;
    }
  }

  const axes: Excel.ChartAxis[] = [];
  if ((defaults.axisDisplay & Axis.Y) !== 0) axes.push(chart.axes.valueAxis);
  if ((defaults.axisDisplay & Axis.X) !== 0) axes.push(chart.axes.categoryAxis);
  if (finishing.tickMarks) {
    for (const axis of axes) {
      axis.majorTickMark = finishing.tickMarks;
      axis.minorTickMark = "None";
    }
  }
  if (finishing.startOnAxis && (defaults.axisDisplay & Axis.X) !== 0) {
    chart.axes.categoryAxis.isBetweenCategories = false;
  }
}

/**
 * Doughnut hole, square plot area and data-label font. `chartWidth`/`chartHeight` are the chart
 * object's size after it was positioned into its band; `hasLegend` reserves the legend strip at the
 * top of a round chart.
 */
export function queueExtraFinishing(
  chart: Excel.Chart,
  finishing: Finishing,
  seriesCount: number,
  chartWidth: number,
  chartHeight: number,
  hasLegend: boolean
): void {
  if (finishing.doughnutHoleSize !== undefined && seriesCount > 0) {
    chart.series.getItemAt(0).doughnutHoleSize = finishing.doughnutHoleSize;
  }
  if (finishing.squarePlot) {
    // Leave room for a legend row at the top (twice the strip height covers the label padding).
    const plot = squarePlotArea(chartWidth, chartHeight, hasLegend ? 2 * legendHeight : 0);
    chart.plotArea.position = "Custom";
    chart.plotArea.left = plot.left;
    chart.plotArea.top = plot.top;
    chart.plotArea.width = plot.width;
    chart.plotArea.height = plot.height;
  }
  if (finishing.dataLabelFont) {
    for (let i = 0; i < seriesCount; i++) {
      const font = chart.series.getItemAt(i).dataLabels.format.font;
      font.name = fontPrimary;
      font.size = axisFontSize;
      font.color = axisFontColor;
    }
  }
}

/** Whether `queueExtraFinishing` has anything to do for this finishing. */
export function hasExtraFinishing(finishing: Finishing): boolean {
  return (
    finishing.doughnutHoleSize !== undefined ||
    finishing.squarePlot === true ||
    finishing.dataLabelFont === true
  );
}
