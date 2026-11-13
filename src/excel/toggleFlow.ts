/* global Excel OfficeExtension */
/**
 * The five chart toggles — ported from `modEngineToggles.bas`. Each acts on the selected chart in
 * one `Excel.run`: read the current state from the workbook (never from stored state, as in the
 * VBA), take one step through its cycle (`logic/toggleCycles`), apply it, and report the new state
 * so the pane can show it.
 *
 * The VBA data-label toggle acted on the series the user had clicked. Office.js cannot see that, so
 * the caller passes the scope: all series, or one series index from the pane's dropdown.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult whose value is a status line.
 */
import { RunResult, runExcel, shapesSupported, fillReadSupported, unsupported } from "./session";
import { styleGridlines, styleLegend } from "./chartStyle";
import { queueExtraFinishing } from "./chartFinishing";
import {
  nextGridlines,
  nextAxisLabels,
  nextAxisTitles,
  nextDataLabels,
  dataLabelStateFrom,
  isCenterOnlyLabelType,
  describeAxisPair,
  AxisPair,
  DataLabelState,
} from "../logic/toggleCycles";
import {
  axesOfChartType,
  colourTargetForChartType,
  isPieChartType,
  legendNotApplicable,
} from "../logic/chartClass";
import { contrastColorForFill } from "../logic/colorContrast";
import { readAxisTitles, setAxisTitles, NOT_STYLED } from "../chrome/chromeRelayout";
import { fontPrimary, axisFontSize, axisFontColor } from "../config/fonts";
import { colorBrand3 } from "../config/brand";

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";
const NO_AXES = "This chart type has no axes to change.";

interface ActiveChart {
  chart: Excel.Chart;
  chartType: string;
  name: string;
  seriesCount: number;
}

async function resolveActive(ctx: Excel.RequestContext): Promise<ActiveChart> {
  const chart = ctx.workbook.getActiveChartOrNullObject();
  chart.load("isNullObject");
  await ctx.sync();
  if (chart.isNullObject) {
    throw new Error(NO_CHART_SELECTED);
  }
  chart.load("chartType, name");
  chart.series.load("count");
  await ctx.sync();
  return {
    chart,
    chartType: chart.chartType,
    name: chart.name,
    seriesCount: chart.series.count,
  };
}

/** Run a toggle with the standard gating. */
function runToggle(fn: (ctx: Excel.RequestContext) => Promise<string>): Promise<RunResult<string>> {
  if (!shapesSupported()) {
    return Promise.resolve(unsupported<string>());
  }
  return runExcel<string>(fn);
}

/** The axes of a chart that exist, as proxies. */
function axisProxies(chart: Excel.Chart, available: AxisPair) {
  return {
    value: available.value ? chart.axes.valueAxis : undefined,
    category: available.category ? chart.axes.categoryAxis : undefined,
  };
}

// --- Gridlines ----------------------------------------------------------------------------

/** Cycle major gridlines: none → value axis → category axis → both. */
export function toggleGridlines(): Promise<RunResult<string>> {
  return runToggle(async (ctx) => {
    const { chart, chartType } = await resolveActive(ctx);
    const available = axesOfChartType(chartType);
    if (!available.value && !available.category) {
      throw new Error(NO_AXES);
    }
    const axes = axisProxies(chart, available);
    axes.value?.majorGridlines.load("visible");
    axes.category?.majorGridlines.load("visible");
    await ctx.sync();

    const current = {
      value: axes.value?.majorGridlines.visible ?? false,
      category: axes.category?.majorGridlines.visible ?? false,
    };
    let next = nextGridlines(current);
    next = { value: next.value && available.value, category: next.category && available.category };
    if (axes.value) styleGridlines(axes.value, next.value);
    if (axes.category) styleGridlines(axes.category, next.category);
    await ctx.sync();
    return `Gridlines: ${describeAxisPair(next, "value axis", "category axis")}`;
  });
}

// --- Axis labels --------------------------------------------------------------------------

/** Cycle axis tick labels: none → category (X) → value (Y) → both. */
export function toggleAxisLabels(): Promise<RunResult<string>> {
  return runToggle(async (ctx) => {
    const { chart, chartType } = await resolveActive(ctx);
    const available = axesOfChartType(chartType);
    if (!available.value && !available.category) {
      throw new Error(NO_AXES);
    }
    const axes = axisProxies(chart, available);
    axes.value?.load("tickLabelPosition, visible");
    axes.category?.load("tickLabelPosition, visible");
    await ctx.sync();

    const shown = (axis?: Excel.ChartAxis) =>
      axis !== undefined && axis.visible && axis.tickLabelPosition !== "None";
    const next = nextAxisLabels({ value: shown(axes.value), category: shown(axes.category) });

    const apply = (axis: Excel.ChartAxis | undefined, show: boolean) => {
      if (!axis) return;
      if (show) {
        axis.visible = true;
        axis.tickLabelPosition = "NextToAxis";
        axis.format.font.name = fontPrimary;
        axis.format.font.size = axisFontSize;
        axis.format.font.color = axisFontColor;
      } else {
        axis.tickLabelPosition = "None";
      }
    };
    apply(axes.value, next.value);
    apply(axes.category, next.category);
    await ctx.sync();
    return `Axis labels: ${describeAxisPair(next, "Y axis", "X axis")}`;
  });
}

// --- Legend -------------------------------------------------------------------------------

/** Show or hide the legend. Round charts re-centre their plot to make room. */
export function toggleLegend(): Promise<RunResult<string>> {
  return runToggle(async (ctx) => {
    const { chart, chartType, seriesCount } = await resolveActive(ctx);
    if (legendNotApplicable(chartType, seriesCount)) {
      throw new Error("This chart has only one data series, or its type has no useful legend.");
    }
    chart.legend.load("visible");
    chart.load("width, height");
    await ctx.sync();

    const show = !chart.legend.visible;
    styleLegend(chart, show);
    await ctx.sync();

    if (isPieChartType(chartType)) {
      // Best-effort: re-centre the square plot below (or without) the legend row.
      try {
        queueExtraFinishing(chart, { squarePlot: true }, seriesCount, chart.width, chart.height, show);
        await ctx.sync();
      } catch {
        // The legend change stands; Excel lays the plot out itself.
      }
    }
    return `Legend: ${show ? "on" : "off"}`;
  });
}

// --- Axis titles --------------------------------------------------------------------------

/** Cycle the axis-title boxes: both → Y → X → none, moving the chart band to match. */
export function toggleAxisTitles(): Promise<RunResult<string>> {
  return runToggle(async (ctx) => {
    const { chart, chartType, name } = await resolveActive(ctx);
    const available = axesOfChartType(chartType);
    if (!available.value && !available.category) {
      throw new Error(NO_AXES);
    }
    const sheet = chart.worksheet;
    const current = await readAxisTitles(ctx, sheet, name);
    if (current === null) {
      throw new Error(NOT_STYLED);
    }
    const next = nextAxisTitles({ value: current.y, category: current.x }, available);
    await setAxisTitles(ctx, sheet, chart, name, { y: next.value, x: next.category });
    return `Axis titles: ${describeAxisPair(next, "Y axis", "X axis")}`;
  });
}

// --- Data labels --------------------------------------------------------------------------

/** All series, or one series by 0-based index. */
export type LabelScope = "all" | number;

function describeLabels(state: DataLabelState): string {
  switch (state) {
    case "outside":
      return "outside end";
    case "inside":
      return "inside centre";
    default:
      return "none";
  }
}

/** Try to give every target series labels at `position`; false when the host refuses it. */
async function trySetLabelPosition(
  ctx: Excel.RequestContext,
  targets: Excel.ChartSeries[],
  position: "OutsideEnd" | "Center"
): Promise<boolean> {
  try {
    for (const series of targets) {
      series.hasDataLabels = true;
      series.dataLabels.position = position;
    }
    await ctx.sync();
    return true;
  } catch {
    return false;
  }
}

/**
 * Colour inside labels for contrast against what they sit on: per point for round charts, per
 * series for filled charts, dark for lines/markers (labels float on the white plot). Reading a fill
 * needs ExcelApi 1.16; without it the labels stay dark.
 */
async function colourInsideLabels(
  ctx: Excel.RequestContext,
  chart: Excel.Chart,
  chartType: string,
  targets: Excel.ChartSeries[]
): Promise<void> {
  const target = colourTargetForChartType(chartType);
  if (target !== "fill" && target !== "point") return;
  if (!fillReadSupported()) return;

  if (target === "point") {
    const points = chart.series.getItemAt(0).points;
    points.load("count");
    await ctx.sync();
    const reads: OfficeExtension.ClientResult<string>[] = [];
    for (let i = 0; i < points.count; i++) {
      reads.push(points.getItemAt(i).format.fill.getSolidColor());
    }
    await ctx.sync();
    reads.forEach((r, i) => {
      if (r.value) {
        points.getItemAt(i).dataLabel.format.font.color = contrastColorForFill(r.value);
      }
    });
    await ctx.sync();
    return;
  }

  const reads = targets.map((s) => s.format.fill.getSolidColor());
  await ctx.sync();
  reads.forEach((r, i) => {
    if (r.value) {
      targets[i].dataLabels.format.font.color = contrastColorForFill(r.value);
    }
  });
  await ctx.sync();
}

/** Cycle data labels: none → outside end → inside centre (none ↔ inside on stacked charts). */
export function toggleDataLabels(scope: LabelScope): Promise<RunResult<string>> {
  return runToggle(async (ctx) => {
    const { chart, chartType, seriesCount } = await resolveActive(ctx);
    if (seriesCount === 0) {
      throw new Error("This chart has no data series.");
    }
    if (scope !== "all" && (scope < 0 || scope >= seriesCount)) {
      throw new Error("That series no longer exists. Pick another one.");
    }
    const indices =
      scope === "all" ? Array.from({ length: seriesCount }, (_, i) => i) : [scope];
    const targets = indices.map((i) => chart.series.getItemAt(i));

    // Current state from the first target series.
    const first = targets[0];
    first.load("hasDataLabels");
    await ctx.sync();
    let position: string | undefined;
    if (first.hasDataLabels) {
      try {
        first.dataLabels.load("position");
        await ctx.sync();
        position = first.dataLabels.position;
      } catch {
        position = undefined; // unreadable: treat as a non-standard position
      }
    }
    const next = nextDataLabels(
      dataLabelStateFrom(first.hasDataLabels, position),
      isCenterOnlyLabelType(chartType)
    );

    if (next === "none") {
      targets.forEach((s) => (s.hasDataLabels = false));
      await ctx.sync();
      return "Data labels: none";
    }

    // Preferred position, falling back to the other one (as the VBA did), then to Excel's default.
    const preferred = next === "outside" ? "OutsideEnd" : "Center";
    const fallback = next === "outside" ? "Center" : "OutsideEnd";
    if (!(await trySetLabelPosition(ctx, targets, preferred))) {
      if (!(await trySetLabelPosition(ctx, targets, fallback))) {
        targets.forEach((s) => (s.hasDataLabels = true));
        await ctx.sync();
      }
    }

    for (const series of targets) {
      const font = series.dataLabels.format.font;
      font.name = fontPrimary;
      font.size = axisFontSize;
      font.color = colorBrand3;
    }
    await ctx.sync();

    if (next === "inside") {
      try {
        await colourInsideLabels(ctx, chart, chartType, targets);
      } catch {
        // Labels stay dark; readable on the light fills, the common case.
      }
    }
    return `Data labels: ${describeLabels(next)}`;
  });
}

/** Names of the selected chart's series, for the data-label scope dropdown. */
export async function listSeriesNames(): Promise<RunResult<string[]>> {
  if (!shapesSupported()) {
    return unsupported<string[]>();
  }
  return runExcel<string[]>(async (ctx) => {
    const chart = ctx.workbook.getActiveChartOrNullObject();
    chart.load("isNullObject");
    await ctx.sync();
    if (chart.isNullObject) {
      return [];
    }
    chart.series.load("items/name");
    await ctx.sync();
    return chart.series.items.map((s) => s.name);
  });
}
