/* global Excel OfficeExtension */
/**
 * Colour-tooling actions on the ACTIVE chart — the task-pane-driven replacement for the VBA ribbon
 * colour buttons + Selection model. The Color Picker is its own task pane (no shared React state
 * with the builder), so it can't be handed a chart name; it acts on the chart the user has selected
 * in Excel, resolved via `getActiveChartOrNullObject` (Excel's own "select then format" model).
 *
 * Each action is one `Excel.run` with explicit sync boundaries: resolve+null-check the active chart,
 * compute the fills with the pure `logic/seriesFills`, then apply via `seriesRecolorer`. The
 * ordering/limits are pure and unit-tested; over-limit throws a RangeError here, which `runExcel`
 * turns into a typed error result (the UI shows the non-blocking "too many series" text).
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { Hex } from "../config/brand";
import type { RampName } from "../config/ramps";
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { paletteFills, singleRampFills, divergingFills, invertFills } from "../logic/seriesFills";
import { isRemoveFillPayload, parseFillPayload, colorFromName } from "../logic/colorFill";
import { applySeriesFills, applyOneSeriesFill, applyAllSeriesFill } from "./seriesRecolorer";

/** How to recolour every series. */
export type RecolourMode =
  | { kind: "palette"; useAltOrder: boolean }
  | { kind: "single"; rampName: RampName }
  | { kind: "diverging"; tag: string }
  | { kind: "invert" };

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";

/**
 * Resolve the active (selected) chart, syncing to null-check it. Throws a friendly error when no
 * chart is selected — `runExcel` turns it into a typed `{ok:false, kind:"error"}` for the UI.
 */
async function resolveActiveChart(ctx: Excel.RequestContext): Promise<Excel.Chart> {
  const active = ctx.workbook.getActiveChartOrNullObject();
  active.load("isNullObject");
  await ctx.sync();
  if (active.isNullObject) {
    throw new Error(NO_CHART_SELECTED);
  }
  return active;
}

/** Recolour all series of the active chart per `mode`. */
export async function recolourSeries(mode: RecolourMode): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  return runExcel<void>(async (ctx) => {
    const chart = await resolveActiveChart(ctx); // Sync A — active-chart null-check.
    chart.series.load("count");

    // Invert needs each series' current fill colour; queue the reads before the first sync.
    let currentColors: OfficeExtension.ClientResult<string>[] = [];
    if (mode.kind === "invert") {
      // Series count isn't known yet, so read it first, then the colours in a second sync.
      await ctx.sync(); // Sync B — series count.
      currentColors = readCurrentFills(chart, chart.series.count);
      await ctx.sync(); // Sync C — current fill colours resolved.
    } else {
      await ctx.sync(); // Sync B — series count.
    }

    const count = chart.series.count;
    const fills = computeFills(mode, count, currentColors);

    applySeriesFills(chart, fills);
    await ctx.sync(); // Sync D — fills applied.
  });
}

/** Queue a `getSolidColor()` read for every series (caller syncs, then reads `.value`). */
function readCurrentFills(
  chart: Excel.Chart,
  count: number
): OfficeExtension.ClientResult<string>[] {
  const reads: OfficeExtension.ClientResult<string>[] = [];
  for (let i = 0; i < count; i++) {
    reads.push(chart.series.getItemAt(i).format.fill.getSolidColor());
  }
  return reads;
}

/** Pure fill-list selection per mode (throws RangeError on over-limit / invalid tag). */
function computeFills(
  mode: RecolourMode,
  count: number,
  currentColors: OfficeExtension.ClientResult<string>[]
): Hex[] {
  switch (mode.kind) {
    case "palette":
      return paletteFills(count, mode.useAltOrder);
    case "single":
      return singleRampFills(mode.rampName, count);
    case "diverging":
      return divergingFills(mode.tag, count);
    case "invert":
      return invertFills(currentColors.map((c) => normalizeHex(c.value)));
  }
}

/** Office returns `#RRGGBB`; normalize to upper-case `#RRGGBB` so re-application is stable. */
function normalizeHex(color: string): Hex {
  const c = color.startsWith("#") ? color : `#${color}`;
  return c.toUpperCase();
}

/**
 * Per-element fill on the active chart: apply `payload` to one series (`target` index) or all series
 * (`"all"`). `payload` is a fill tag like `"DATA1"` or `"NONE"`. Transparency in the payload is
 * parsed but NOT applied — Office.js chart-series fills have no transparency control (solid only).
 */
export async function applyElementFill(
  target: number | "all",
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
    const chart = await resolveActiveChart(ctx); // Sync A — active-chart null-check.
    chart.series.load("count");
    await ctx.sync(); // Sync B — series count.

    if (target === "all") {
      applyAllSeriesFill(chart, chart.series.count, hex);
    } else {
      applyOneSeriesFill(chart, target, hex);
    }
    await ctx.sync(); // Sync C — fill applied.
  });
}

/** List the active chart's series names so the UI can populate the element selector. */
export async function listSeries(): Promise<RunResult<string[]>> {
  if (!shapesSupported()) {
    return unsupported<string[]>();
  }

  return runExcel<string[]>(async (ctx) => {
    const chart = await resolveActiveChart(ctx); // Sync A — active-chart null-check.
    chart.series.load("items/name");
    await ctx.sync(); // Sync B — series names.
    return chart.series.items.map((s) => s.name);
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
