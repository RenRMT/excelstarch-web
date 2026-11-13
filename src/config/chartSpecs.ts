/**
 * One description per chart kind — the single place that says which Excel chart type a kind maps
 * to, where its colour lives, and which formatting defaults it uses. The interop layer reads this
 * table instead of keeping its own switches.
 *
 * `excelType` is the Office.js `Excel.ChartType` string value, written as a literal so this module
 * stays free of `Excel` imports and can be unit-tested.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { ChartDefaults } from "./chartDefaults";
import {
  barChartDefaults,
  columnChartDefaults,
  lineChartDefaults,
  areaChartDefaults,
  scatterChartDefaults,
  pieChartDefaults,
  treemapChartDefaults,
  boxWhiskerChartDefaults,
} from "./chartDefaults";

/** The chart kinds the builder can create. */
export type ChartKind =
  | "bar"
  | "column"
  | "line"
  | "area"
  | "scatter"
  | "pie"
  | "treemap"
  | "boxwhisker";

/**
 * Where a chart's series colour lives:
 *   - `fill`       the series fill (bar, column, area, box & whisker)
 *   - `line`       the series line (line charts without markers)
 *   - `lineMarker` the line and its markers
 *   - `marker`     the markers only (scatter)
 *   - `point`      each data point of the first series (pie slices, treemap tiles)
 */
export type ColourTarget = "fill" | "line" | "lineMarker" | "marker" | "point";

export interface ChartSpec {
  /** Office.js `Excel.ChartType` string value. */
  excelType: string;
  /** Button label in the builder. */
  label: string;
  colourTarget: ColourTarget;
  defaults: () => ChartDefaults;
}

export const chartSpecs: Readonly<Record<ChartKind, ChartSpec>> = {
  bar: {
    excelType: "BarClustered",
    label: "Bar",
    colourTarget: "fill",
    defaults: barChartDefaults,
  },
  column: {
    excelType: "ColumnClustered",
    label: "Column",
    colourTarget: "fill",
    defaults: columnChartDefaults,
  },
  line: {
    excelType: "Line",
    label: "Line",
    colourTarget: "line",
    defaults: lineChartDefaults,
  },
  area: {
    excelType: "Area",
    label: "Area",
    colourTarget: "fill",
    defaults: areaChartDefaults,
  },
  scatter: {
    excelType: "XYScatter",
    label: "Scatter",
    colourTarget: "marker",
    defaults: scatterChartDefaults,
  },
  pie: {
    excelType: "Pie",
    label: "Pie",
    colourTarget: "point",
    defaults: pieChartDefaults,
  },
  treemap: {
    excelType: "Treemap",
    label: "Treemap",
    colourTarget: "point",
    defaults: treemapChartDefaults,
  },
  boxwhisker: {
    excelType: "Boxwhisker",
    label: "Box & Whisker",
    colourTarget: "fill",
    defaults: boxWhiskerChartDefaults,
  },
};

/** The chart kinds in builder display order. */
export const chartKinds: readonly ChartKind[] = Object.keys(chartSpecs) as ChartKind[];
