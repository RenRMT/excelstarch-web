/* global Excel OfficeExtension */
/**
 * Colour-tooling actions on the ACTIVE chart — the task-pane-driven replacement for the VBA ribbon
 * colour buttons + Selection model. The Color Picker is its own task pane (no shared React state
 * with the builder), so it can't be handed a chart name; it acts on the chart the user has selected
 * in Excel, resolved via `getActiveChartOrNullObject` (Excel's own "select then format" model).
 *
 * The chart can be of any type, including charts this add-in did not create, so each action first
 * classifies it (`logic/chartClass`) to find where its colour lives: the series fill, the series
 * line, the markers, or — for pie/treemap — the points of the first series. The "elements" being
 * coloured are therefore the series, or the points for a point-coloured chart.
 *
 * Each action is one `Excel.run` with explicit sync boundaries. The ordering/limits are pure and
 * unit-tested; over-limit throws a RangeError here, which `runExcel` turns into a typed error
 * result (the UI shows the non-blocking "too many series" text).
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { Hex } from "../config/brand";
import type { RampName } from "../config/ramps";
import type { ColourTarget } from "../config/chartSpecs";
import { colourTargetForChartType } from "../logic/chartClass";
import { RunResult, runExcel, shapesSupported, fillReadSupported, unsupported } from "./session";
import { paletteFills, singleRampFills, divergingFills, invertFills } from "../logic/seriesFills";
import { isRemoveFillPayload, parseFillPayload, colorFromName } from "../logic/colorFill";
import { applyColours, applyOneColour } from "./seriesRecolorer";

/** How to recolour every element. */
export type RecolourMode =
  | { kind: "palette"; useAltOrder: boolean }
  | { kind: "single"; rampName: RampName }
  | { kind: "diverging"; tag: string }
  | { kind: "invert" };

/** What the Color Picker's element selector lists for the active chart. */
export interface ChartElements {
  /** `point` when the chart is coloured per point (pie, treemap); otherwise `series`. */
  kind: "series" | "point";
  names: string[];
}

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";
const INVERT_NEEDS_NEWER_EXCEL =
  "Inverting the colours of this chart type needs a newer Excel (ExcelApi 1.16). " +
  "Apply a ramp or the palette instead.";

interface ColourContext {
  chart: Excel.Chart;
  target: ColourTarget;
  /** Number of elements to colour: series, or points of the first series for a point target. */
  count: number;
}

/**
 * Resolve the active chart, classify it and count its colourable elements. Throws a friendly
 * error when no chart is selected — `runExcel` turns it into a typed `{ok:false, kind:"error"}`.
 */
async function resolveColourContext(ctx: Excel.RequestContext): Promise<ColourContext> {
  const chart = ctx.workbook.getActiveChartOrNullObject();
  chart.load("isNullObject");
  await ctx.sync(); // active-chart null-check.
  if (chart.isNullObject) {
    throw new Error(NO_CHART_SELECTED);
  }

  chart.load("chartType");
  chart.series.load("count");
  await ctx.sync(); // chart type + series count.

  const target = colourTargetForChartType(chart.chartType);
  if (target !== "point" || chart.series.count === 0) {
    return { chart, target, count: chart.series.count };
  }

  const points = chart.series.getItemAt(0).points;
  points.load("count");
  await ctx.sync(); // point count.
  return { chart, target, count: points.count };
}

/** Recolour all elements of the active chart per `mode`. */
export async function recolourSeries(mode: RecolourMode): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  return runExcel<void>(async (ctx) => {
    const { chart, target, count } = await resolveColourContext(ctx);

    let fills: Hex[];
    if (mode.kind === "invert") {
      fills = invertFills(await readCurrentColours(ctx, chart, target, count));
    } else {
      fills = computeFills(mode, count);
    }

    applyColours(chart, target, fills);
    await ctx.sync(); // colours applied.
  });
}

/**
 * Read the current colour of every element from wherever this chart type keeps it. Fill colours
 * can only be read back on ExcelApi 1.16+; line and marker colours are plain properties.
 */
async function readCurrentColours(
  ctx: Excel.RequestContext,
  chart: Excel.Chart,
  target: ColourTarget,
  count: number
): Promise<Hex[]> {
  if (target === "line" || target === "lineMarker") {
    const lines = range(count).map((i) => chart.series.getItemAt(i).format.line);
    lines.forEach((line) => line.load("color"));
    await ctx.sync();
    return lines.map((line) => normalizeHex(line.color));
  }

  if (target === "marker") {
    const series = range(count).map((i) => chart.series.getItemAt(i));
    series.forEach((s) => s.load("markerBackgroundColor"));
    await ctx.sync();
    return series.map((s) => normalizeHex(s.markerBackgroundColor));
  }

  if (!fillReadSupported()) {
    throw new Error(INVERT_NEEDS_NEWER_EXCEL);
  }
  let reads: OfficeExtension.ClientResult<string>[];
  if (target === "point") {
    const points = chart.series.getItemAt(0).points;
    reads = range(count).map((i) => points.getItemAt(i).format.fill.getSolidColor());
  } else {
    reads = range(count).map((i) => chart.series.getItemAt(i).format.fill.getSolidColor());
  }
  await ctx.sync();
  return reads.map((r) => normalizeHex(r.value));
}

function range(count: number): number[] {
  return Array.from({ length: count }, (_, i) => i);
}

/** Pure fill-list selection per mode (throws RangeError on over-limit / invalid tag). */
function computeFills(mode: Exclude<RecolourMode, { kind: "invert" }>, count: number): Hex[] {
  switch (mode.kind) {
    case "palette":
      return paletteFills(count, mode.useAltOrder);
    case "single":
      return singleRampFills(mode.rampName, count);
    case "diverging":
      return divergingFills(mode.tag, count);
  }
}

/** Office returns `#RRGGBB`; normalize to upper-case `#RRGGBB` so re-application is stable. */
function normalizeHex(color: string): Hex {
  const c = color.startsWith("#") ? color : `#${color}`;
  return c.toUpperCase();
}

/**
 * Per-element colour on the active chart: apply `payload` to one element (`element` index) or all
 * elements (`"all"`). `payload` is a fill tag like `"DATA1"` or `"NONE"`. Transparency in the
 * payload is parsed but NOT applied — Office.js chart fills have no transparency control.
 */
export async function applyElementFill(
  element: number | "all",
  payload: string
): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  // Resolve the payload to a hex (or null for remove) before touching the host.
  let hex: Hex | null;
  if (isRemoveFillPayload(payload)) {
    hex = null;
  } else {
    const { name } = parseFillPayload(payload);
    hex = colorFromName(name);
    if (hex === null) {
      return { ok: false, kind: "error", message: `Unknown fill colour "${name}".` };
    }
  }

  return runExcel<void>(async (ctx) => {
    const { chart, target, count } = await resolveColourContext(ctx);

    if (element === "all") {
      applyColours(
        chart,
        target,
        range(count).map(() => hex)
      );
    } else {
      applyOneColour(chart, target, element, hex);
    }
    await ctx.sync(); // colour applied.
  });
}

/**
 * List the active chart's colourable elements so the UI can populate the element selector: the
 * series names, or — for a point-coloured chart — the category names of the first series' points.
 */
export async function listElements(): Promise<RunResult<ChartElements>> {
  if (!shapesSupported()) {
    return unsupported<ChartElements>();
  }

  return runExcel<ChartElements>(async (ctx) => {
    const { chart, target, count } = await resolveColourContext(ctx);

    if (target !== "point") {
      chart.series.load("items/name");
      await ctx.sync(); // series names.
      return { kind: "series", names: chart.series.items.map((s) => s.name) };
    }

    // Point names are the category labels. Reading them can fail on some chart types, so it gets
    // its own guarded sync and falls back to numbered points.
    let names = range(count).map((i) => `Point ${i + 1}`);
    if (count > 0) {
      try {
        const categories = chart.series.getItemAt(0).getDimensionValues("Categories");
        await ctx.sync();
        if (categories.value.length === count) {
          names = categories.value.map((c, i) => (c ? String(c) : names[i]));
        }
      } catch {
        // Keep the numbered fallback.
      }
    }
    return { kind: "point", names };
  });
}

/** Whether a chart is currently selected, so the Color Picker can enable/disable its controls. */
export async function hasActiveChart(): Promise<RunResult<boolean>> {
  if (!shapesSupported()) {
    return unsupported<boolean>();
  }

  return runExcel<boolean>(async (ctx) => {
    const active = ctx.workbook.getActiveChartOrNullObject();
    active.load("isNullObject");
    await ctx.sync();
    return !active.isNullObject;
  });
}
