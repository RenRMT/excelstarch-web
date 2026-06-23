/* global Excel */
/**
 * Non-shape chart formatting — the part of `modEngineBuilder`/`ApplyDefaultFormatting` that lives
 * on the chart object (gridlines, axis visibility, axis/gridline line styling, legend), as opposed
 * to the chrome text/logo (which are worksheet shapes in chrome/).
 *
 * Driven by the per-type `ChartDefaults`: bar shows X-axis gridlines, column shows Y-axis
 * gridlines; both show both axes and hide the legend.
 *
 * QUEUES property writes only — the caller syncs. No `context.sync()` here.
 *
 * INTEROP: Office.js chart API.
 */
import type { ChartDefaults } from "../config/chartDefaults";
import { Axis } from "../config/enums";
import { colorNeutral2, colorWhite } from "../config/brand";
import { gridlineWeight, axisLineWeight } from "../config/geometry";
import { fontPrimary, axisFontSize, axisFontColor } from "../config/fonts";

const GRIDLINE_COLOR = colorNeutral2; // light grey value-scale lines (matches the VBA)

export function applyChartStyle(chart: Excel.Chart, defaults: ChartDefaults): void {
  const axes = chart.axes;

  // --- Suppress the chart's own title; the branded title lives in the chrome title shape ---
  chart.title.visible = false;

  // --- Transparent, borderless chart area so only the white chrome canvas shows behind it ---
  chart.format.fill.clear();
  chart.format.border.clear();

  // --- Gridlines: only on the axes named by defaults.gridlines ---
  const showYGrid = (defaults.gridlines & Axis.Y) !== 0;
  const showXGrid = (defaults.gridlines & Axis.X) !== 0;
  styleGridlines(axes.valueAxis, showYGrid);
  styleGridlines(axes.categoryAxis, showXGrid);

  // --- Axis visibility ---
  axes.valueAxis.visible = (defaults.axisDisplay & Axis.Y) !== 0;
  axes.categoryAxis.visible = (defaults.axisDisplay & Axis.X) !== 0;

  // --- Axis line + tick-label font (white axis lines; branded labels) ---
  styleAxisLineAndFont(axes.valueAxis);
  styleAxisLineAndFont(axes.categoryAxis);

  // --- Legend ---
  chart.legend.visible = defaults.legend;
}

function styleGridlines(axis: Excel.ChartAxis, visible: boolean): void {
  axis.majorGridlines.visible = visible;
  if (visible) {
    const line = axis.majorGridlines.format.line;
    line.color = GRIDLINE_COLOR;
    line.weight = gridlineWeight;
  }
}

function styleAxisLineAndFont(axis: Excel.ChartAxis): void {
  axis.format.line.color = colorWhite;
  axis.format.line.weight = axisLineWeight;
  const font = axis.format.font;
  font.name = fontPrimary;
  font.size = axisFontSize;
  font.color = axisFontColor;
}
