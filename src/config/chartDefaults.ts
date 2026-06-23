/**
 * Per-chart-type formatting defaults — ported from `modConfigDerived.bas`. Bundles the
 * formatting options the chart pipeline consumes into one object per chart type.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import { Axis, defaultGridlines, defaultAxisDisplay, defaultLegend } from "./enums";

/** Formatting options for a new or reformatted chart. */
export interface ChartDefaults {
  /** Which axes show gridlines. */
  gridlines: Axis;
  /** Which axes are displayed (HasAxis). */
  axisDisplay: Axis;
  /** Whether to show the legend. */
  legend: boolean;
  /**
   * Chartex only: add a worksheet Y-axis title box (box & whisker has a value axis).
   * Classic factories leave this false and ignore it.
   */
  showYAxisTitle: boolean;
}

/** Global defaults — minimal chrome. */
export function defaultChartDefaults(): ChartDefaults {
  return {
    gridlines: defaultGridlines,
    axisDisplay: defaultAxisDisplay,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function lineChartDefaults(): ChartDefaults {
  return {
    gridlines: Axis.Y,
    axisDisplay: Axis.Both,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function barChartDefaults(): ChartDefaults {
  return {
    // Gridlines on the VALUE axis. A horizontal bar rotates the axes so the value axis runs
    // horizontally — value-axis gridlines therefore render as the expected vertical lines. (Naming
    // the category axis here, as Axis.X once did, would draw horizontal gridlines on a bar chart.)
    gridlines: Axis.Y,
    axisDisplay: Axis.Both,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function columnChartDefaults(): ChartDefaults {
  return {
    gridlines: Axis.Y,
    axisDisplay: Axis.Both,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function areaChartDefaults(): ChartDefaults {
  return {
    gridlines: Axis.Y,
    axisDisplay: Axis.Both,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function scatterChartDefaults(): ChartDefaults {
  return {
    gridlines: Axis.Both,
    axisDisplay: Axis.Both,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function pieChartDefaults(): ChartDefaults {
  return { gridlines: Axis.None, axisDisplay: Axis.None, legend: true, showYAxisTitle: false };
}

export function treemapChartDefaults(): ChartDefaults {
  return {
    gridlines: Axis.None,
    axisDisplay: Axis.None,
    legend: defaultLegend,
    showYAxisTitle: false,
  };
}

export function boxWhiskerChartDefaults(): ChartDefaults {
  return { gridlines: Axis.Y, axisDisplay: Axis.Both, legend: defaultLegend, showYAxisTitle: true };
}
