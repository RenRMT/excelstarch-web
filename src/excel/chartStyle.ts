/* global Excel */
/**
 * Non-shape chart formatting — the part of `modEngineBuilder`/`ApplyDefaultFormatting` that lives
 * on the chart object (gridlines, axis visibility, axis/gridline line styling, legend), as opposed
 * to the chrome text (which are worksheet shapes in chrome/).
 *
 * Driven by the per-type `ChartDefaults`: e.g. column shows Y-axis gridlines and both axes; pie and
 * treemap display no axes (and treemap no legend). Axis styling is gated on `axisDisplay` so a
 * no-axis type never touches a non-existent value/category axis.
 *
 * QUEUES property writes only — the caller syncs. No `context.sync()` here.
 *
 * INTEROP: Office.js chart API.
 */
import type { ChartDefaults } from "../config/chartDefaults";
import { Axis } from "../config/enums";
import { colorNeutral2, colorWhite } from "../config/brand";
import { gridlineWeight, axisLineWeight } from "../config/geometry";
import { fontPrimary, axisFontSize, axisFontColor, legendFontColor } from "../config/fonts";

const GRIDLINE_COLOR = colorNeutral2; // light grey value-scale lines (matches the VBA)

/**
 * `keepLegend` leaves the legend as it is — used when styling a chart type the add-in knows nothing
 * about (Apply Chart Style's "other" class); such a profile also displays no axes, so the axes are
 * left alone too.
 */
export function applyChartStyle(
  chart: Excel.Chart,
  defaults: ChartDefaults,
  keepLegend = false
): void {
  const axes = chart.axes;

  // --- Suppress the chart's own title; the branded title lives in the chrome title shape ---
  chart.title.visible = false;

  // --- Transparent, borderless chart area so only the white chrome canvas shows behind it ---
  chart.format.fill.clear();
  chart.format.border.clear();

  // --- Chart-wide font, so the legend and data labels match the axes and chrome text ---
  chart.format.font.name = fontPrimary;

  // --- Axes: touch an axis ONLY when this chart type displays it. A treemap (and pie) have no
  // value/category axes, so reading axes.valueAxis/categoryAxis there would queue writes against a
  // non-existent axis and throw on sync. Gating on axisDisplay keeps the no-axis types safe. ---
  if ((defaults.axisDisplay & Axis.Y) !== 0) {
    styleGridlines(axes.valueAxis, (defaults.gridlines & Axis.Y) !== 0);
    axes.valueAxis.visible = true;
    styleAxisLineAndFont(axes.valueAxis);
  }
  if ((defaults.axisDisplay & Axis.X) !== 0) {
    styleGridlines(axes.categoryAxis, (defaults.gridlines & Axis.X) !== 0);
    axes.categoryAxis.visible = true;
    styleAxisLineAndFont(axes.categoryAxis);
  }

  // --- Legend (valid on every chart type, including treemap/pie) ---
  if (!keepLegend) {
    styleLegend(chart, defaults.legend);
  }
}

/**
 * Show or hide the legend. When shown it sits in a row at the top of the chart object, in the axis
 * font size and colour (the VBA placed it in a strip under the subtitle; here the chart object
 * starts under the title block, so the top of the chart is that same place).
 */
export function styleLegend(chart: Excel.Chart, visible: boolean): void {
  chart.legend.visible = visible;
  if (visible) {
    chart.legend.position = "Top";
    const font = chart.legend.format.font;
    font.size = axisFontSize;
    font.color = legendFontColor;
  }
}

export function styleGridlines(axis: Excel.ChartAxis, visible: boolean): void {
  axis.majorGridlines.visible = visible;
  if (visible) {
    const line = axis.majorGridlines.format.line;
    line.color = GRIDLINE_COLOR;
    line.weight = gridlineWeight;
  }
}

export function styleAxisLineAndFont(axis: Excel.ChartAxis): void {
  axis.format.line.color = colorWhite;
  axis.format.line.weight = axisLineWeight;
  const font = axis.format.font;
  font.name = fontPrimary;
  font.size = axisFontSize;
  font.color = axisFontColor;
}
