/**
 * One description per chart kind — the single place that says which Excel chart type a kind maps
 * to, where its colour lives, which formatting defaults it uses, and the per-type finishing the VBA
 * builders applied after the shared pipeline (gap width, tick marks, markers, …). The interop layer
 * reads this table instead of keeping its own switches.
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
import {
  seriesGapWidth,
  seriesOverlap,
  stackedOverlap,
  scatterMarkerSize,
  scatterMarkerStyles,
  donutHoleSize,
} from "./geometry";

/** The chart kinds the builder can create. */
export type ChartKind =
  | "column"
  | "columnStacked"
  | "columnStacked100"
  | "bar"
  | "barStacked"
  | "barStacked100"
  | "line"
  | "area"
  | "area100"
  | "pie"
  | "doughnut"
  | "scatter"
  | "bubble"
  | "treemap"
  | "boxwhisker";

/**
 * Where a chart's series colour lives:
 *   - `fill`       the series fill (bar, column, area, bubble, box & whisker)
 *   - `line`       the series line (line charts without markers)
 *   - `lineMarker` the line and its markers
 *   - `marker`     the markers only (scatter)
 *   - `point`      each data point of the first series (pie/doughnut slices, treemap tiles)
 */
export type ColourTarget = "fill" | "line" | "lineMarker" | "marker" | "point";

/** Builder button groups, mirroring the VBA ribbon groups. */
export type ChartFamily = "Column" | "Bar" | "Line & Area" | "Pie" | "Scatter" | "Complex";

/**
 * Per-type touches applied after the shared styling — ported from the per-type builders
 * (`modChartBar`, `modChartLine`, `modChartScatter`, `modChartPie`, …). Every field is optional;
 * an absent field leaves Excel's default.
 */
export interface Finishing {
  /** Space between bars/columns, as a percentage of their width. */
  gapWidth?: number;
  /** Overlap between series, −100..100 (100 for stacked: segments must be flush). */
  overlap?: number;
  /** Major tick marks on both displayed axes; minor tick marks are always switched off. */
  tickMarks?: "None" | "Outside";
  /** Start the category axis on the first data point, so lines/areas run edge to edge. */
  startOnAxis?: boolean;
  /** Remove the default series shadow. */
  noShadow?: boolean;
  /** Give each series a distinct marker shape (cycling) at this size. */
  markers?: { styles: readonly string[]; size: number };
  /** Doughnut hole size in percent. */
  doughnutHoleSize?: number;
  /** Size the plot area as a centred square (round charts). */
  squarePlot?: boolean;
  /** Put data labels in the house font (types that show labels by default). */
  dataLabelFont?: boolean;
}

export interface ChartSpec {
  /** Office.js `Excel.ChartType` string value. */
  excelType: string;
  /** Button label in the builder. */
  label: string;
  family: ChartFamily;
  colourTarget: ColourTarget;
  defaults: () => ChartDefaults;
  finishing: Finishing;
}

const clusteredBars: Finishing = {
  gapWidth: seriesGapWidth,
  overlap: seriesOverlap,
  tickMarks: "None",
  noShadow: true,
};
const stackedBars: Finishing = { ...clusteredBars, overlap: stackedOverlap };
const areas: Finishing = { tickMarks: "None", startOnAxis: true };
const roundCharts: Finishing = { squarePlot: true };

export const chartSpecs: Readonly<Record<ChartKind, ChartSpec>> = {
  column: {
    excelType: "ColumnClustered",
    label: "Column",
    family: "Column",
    colourTarget: "fill",
    defaults: columnChartDefaults,
    finishing: clusteredBars,
  },
  columnStacked: {
    excelType: "ColumnStacked",
    label: "Stacked column",
    family: "Column",
    colourTarget: "fill",
    defaults: columnChartDefaults,
    finishing: stackedBars,
  },
  columnStacked100: {
    excelType: "ColumnStacked100",
    label: "100% column",
    family: "Column",
    colourTarget: "fill",
    defaults: columnChartDefaults,
    finishing: stackedBars,
  },
  bar: {
    excelType: "BarClustered",
    label: "Bar",
    family: "Bar",
    colourTarget: "fill",
    defaults: barChartDefaults,
    finishing: clusteredBars,
  },
  barStacked: {
    excelType: "BarStacked",
    label: "Stacked bar",
    family: "Bar",
    colourTarget: "fill",
    defaults: barChartDefaults,
    finishing: stackedBars,
  },
  barStacked100: {
    excelType: "BarStacked100",
    label: "100% bar",
    family: "Bar",
    colourTarget: "fill",
    defaults: barChartDefaults,
    finishing: stackedBars,
  },
  line: {
    excelType: "Line",
    label: "Line",
    family: "Line & Area",
    colourTarget: "line",
    defaults: lineChartDefaults,
    finishing: { tickMarks: "Outside", startOnAxis: true, noShadow: true },
  },
  area: {
    excelType: "AreaStacked",
    label: "Area",
    family: "Line & Area",
    colourTarget: "fill",
    defaults: areaChartDefaults,
    finishing: areas,
  },
  area100: {
    excelType: "AreaStacked100",
    label: "100% area",
    family: "Line & Area",
    colourTarget: "fill",
    defaults: areaChartDefaults,
    finishing: areas,
  },
  pie: {
    excelType: "Pie",
    label: "Pie",
    family: "Pie",
    colourTarget: "point",
    defaults: pieChartDefaults,
    finishing: roundCharts,
  },
  doughnut: {
    excelType: "Doughnut",
    label: "Donut",
    family: "Pie",
    colourTarget: "point",
    defaults: pieChartDefaults,
    finishing: { ...roundCharts, doughnutHoleSize: donutHoleSize },
  },
  scatter: {
    excelType: "XYScatter",
    label: "Scatter",
    family: "Scatter",
    colourTarget: "marker",
    defaults: scatterChartDefaults,
    finishing: {
      tickMarks: "Outside",
      markers: { styles: scatterMarkerStyles, size: scatterMarkerSize },
    },
  },
  bubble: {
    excelType: "Bubble",
    label: "Bubble",
    family: "Scatter",
    colourTarget: "fill",
    defaults: scatterChartDefaults,
    finishing: { tickMarks: "Outside" },
  },
  treemap: {
    excelType: "Treemap",
    label: "Treemap",
    family: "Complex",
    colourTarget: "point",
    defaults: treemapChartDefaults,
    finishing: { dataLabelFont: true },
  },
  boxwhisker: {
    excelType: "Boxwhisker",
    label: "Box & Whisker",
    family: "Complex",
    colourTarget: "fill",
    defaults: boxWhiskerChartDefaults,
    finishing: { dataLabelFont: true },
  },
};

/** The chart kinds in builder display order. */
export const chartKinds: readonly ChartKind[] = Object.keys(chartSpecs) as ChartKind[];

/** The builder's families in display order, each with its kinds (the first is the default). */
export const chartFamilies: readonly { family: ChartFamily; kinds: ChartKind[] }[] =
  chartKinds.reduce<{ family: ChartFamily; kinds: ChartKind[] }[]>((groups, kind) => {
    const family = chartSpecs[kind].family;
    const group = groups.find((g) => g.family === family);
    if (group) {
      group.kinds.push(kind);
    } else {
      groups.push({ family, kinds: [kind] });
    }
    return groups;
  }, []);

/** The kind whose Excel chart type is `excelType`, if the builder has one. */
export function kindForExcelType(excelType: string): ChartKind | undefined {
  return chartKinds.find((k) => chartSpecs[k].excelType === excelType);
}
