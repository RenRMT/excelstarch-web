/* global Excel */
/**
 * The end-to-end "make this chart branded" actions, each inside ONE `Excel.run` with explicit sync
 * boundaries. This is the orchestrator the task pane calls.
 *
 *   - `createBrandedChart(kind)` — create a chart of `kind` from the selected range, or retype the
 *     selected chart to `kind` (the VBA chart buttons / `GetTargetChart`).
 *   - `styleSelectedChart()` — style the selected chart in place without changing its type, as far
 *     as its type allows (the VBA `ApplyChartStyle`).
 *
 * Both end in `brandChart`, driven by a `StyleProfile` (logic/styleProfile). Re-runs snap the chrome
 * back to its existing canvas origin and keep the text the user typed into the chrome boxes.
 *
 * Sync structure (never sync inside a loop):
 *   A  — read whether a chart is selected (create vs restyle decision).
 *   B  — create/retype the chart; load its name, type, position, parent worksheet, series count.
 *   C  — chrome snapshot (canvas origin, prior group, existing text) + point count for pie/treemap.
 *   D  — reposition + style + colour the chart AND build all chrome shapes (they now exist).
 *   E  — group the chrome shapes (guarded; must follow the shapes-exist sync). The live chart can't
 *        be a group member (disjoint Chart/Shape object models); it's sized into the canvas band.
 *   F… — guarded follow-ups that some chart types or hosts reject: per-type finishing, extra
 *        finishing, and best-effort colouring for unknown types. Each failure costs only itself.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { ChartKind } from "../logic/chartType";
import { profileForKind, profileForChartType, StyleProfile } from "../logic/styleProfile";
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { createChart, resolveChartRange, retypeChart, positionChartIntoBand } from "./chartFactory";
import { applyChartStyle } from "./chartStyle";
import { colorByPalette } from "./seriesColorer";
import {
  queueSeriesAndAxisFinishing,
  queueExtraFinishing,
  hasExtraFinishing,
} from "./chartFinishing";
import { buildChrome } from "../chrome/chromeBuilder";
import { groupChrome, chromeGroupName } from "../chrome/chromeGroup";
import { readChromeSnapshot } from "../chrome/chromeSnapshot";
import { chartBandPosition } from "../chrome/chromeLayout";
import { Axis } from "../config/enums";

export interface CreateChartResult {
  chartName: string;
  /** Id of the worksheet the chart lives on (chart names are only unique per sheet). */
  worksheetId: string;
  groupName: string;
  warnings: string[];
}

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";

/** Create a branded chart of `kind`, or retype + restyle the selected chart to it. */
export async function createBrandedChart(kind: ChartKind): Promise<RunResult<CreateChartResult>> {
  if (!shapesSupported()) {
    return unsupported<CreateChartResult>();
  }

  return runExcel<CreateChartResult>(async (ctx) => {
    // Decide create vs restyle: is a chart currently selected? Also read the selected range's
    // cellCount so a create can expand a single selected cell to its surrounding region.
    const active = ctx.workbook.getActiveChartOrNullObject();
    active.load("isNullObject");
    const selected = ctx.workbook.getSelectedRange();
    selected.load("cellCount");
    await ctx.sync(); // Sync A — restyle-or-create decision + selection size.

    let chart: Excel.Chart;
    if (active.isNullObject) {
      const range = resolveChartRange(selected);
      chart = createChart(ctx.workbook.worksheets.getActiveWorksheet(), range, kind);
    } else {
      chart = active;
      retypeChart(chart, kind);
    }
    return brandChart(ctx, chart, () => profileForKind(kind));
  });
}

/** Style the selected chart in place, keeping its chart type (the VBA "Apply Chart Style"). */
export async function styleSelectedChart(): Promise<RunResult<CreateChartResult>> {
  if (!shapesSupported()) {
    return unsupported<CreateChartResult>();
  }

  return runExcel<CreateChartResult>(async (ctx) => {
    const chart = ctx.workbook.getActiveChartOrNullObject();
    chart.load("isNullObject");
    await ctx.sync(); // Sync A — is a chart selected?
    if (chart.isNullObject) {
      throw new Error(NO_CHART_SELECTED);
    }
    return brandChart(ctx, chart, (chartType) => profileForChartType(chartType));
  });
}

/** Run `queue`, then sync; on failure note `warning` (if any) and carry on. */
async function guardedSync(
  ctx: Excel.RequestContext,
  queue: () => void,
  warnings: string[],
  warning?: string
): Promise<void> {
  try {
    queue();
    await ctx.sync();
  } catch {
    if (warning) {
      warnings.push(warning);
    }
  }
}

/**
 * Style `chart` (already created or retyped, not yet synced) and build its chrome. `profileFor`
 * receives the chart's Excel type once it has been read.
 */
async function brandChart(
  ctx: Excel.RequestContext,
  chart: Excel.Chart,
  profileFor: (chartType: string) => StyleProfile
): Promise<CreateChartResult> {
  // Resolve the chart's OWN worksheet (a selected chart may live on a non-active sheet); the chrome
  // must be built on that sheet.
  const sheet = chart.worksheet;
  sheet.load("id");
  chart.load("name, chartType, left, top");
  chart.series.load("count");
  await ctx.sync(); // Sync B — name, type, position, series count, sheet.

  const chartName = chart.name;
  const seriesCount = chart.series.count;
  const profile = profileFor(chart.chartType);
  const { defaults, colourTarget, finishing } = profile;

  // Pie/doughnut and treemap have a single series whose slices/tiles are POINTS — colour per point,
  // not per series (a solid series fill would flatten the whole chart to one colour).
  const colourByPoint = colourTarget === "point" && seriesCount > 0;
  const firstSeriesPoints = colourByPoint ? chart.series.getItemAt(0).points : undefined;
  firstSeriesPoints?.load("count");

  // Existing chrome: canvas origin (so a re-run snaps back), prior group (so it can be safely
  // ungrouped) and the text already typed into the boxes (so a re-run keeps it).
  const snapshot = await readChromeSnapshot(ctx, sheet, chartName); // Sync C.

  const baseLeft = snapshot.canvasOrigin ? snapshot.canvasOrigin.left : chart.left;
  const baseTop = snapshot.canvasOrigin ? snapshot.canvasOrigin.top : chart.top;
  const colourCount = firstSeriesPoints ? firstSeriesPoints.count : seriesCount;
  const showY = (defaults.axisDisplay & Axis.Y) !== 0;
  const showX = (defaults.axisDisplay & Axis.X) !== 0;

  // Queue all core mutations (no sync between them).
  positionChartIntoBand(chart, baseLeft, baseTop, showY, showX);
  applyChartStyle(chart, defaults, profile.keepAxesAndLegend);
  if (!profile.bestEffortColour) {
    colorByPalette(chart, colourTarget, colourCount);
  }
  const { createdShapeNames, warnings } = buildChrome(
    sheet,
    chartName,
    baseLeft,
    baseTop,
    defaults,
    snapshot.groupExists,
    snapshot.texts
  );
  await ctx.sync(); // Sync D — chart repositioned/styled/coloured AND chrome shapes exist.

  // Group the chrome shapes. Guarded so a rare failure degrades to a warning rather than
  // discarding an otherwise-complete chart + chrome.
  await guardedSync(
    ctx,
    () => groupChrome(sheet, chartName, createdShapeNames),
    warnings,
    "The chart and its chrome were created but the chrome could not be grouped."
  ); // Sync E.

  // Follow-ups that not every chart type or host accepts, each in its own guarded sync.
  if (profile.bestEffortColour) {
    // Unknown types may reject standard colouring; skip silently, as the VBA did.
    await guardedSync(ctx, () => colorByPalette(chart, colourTarget, colourCount), warnings);
  }
  await guardedSync(
    ctx,
    () => queueSeriesAndAxisFinishing(chart, finishing, defaults, seriesCount),
    warnings,
    "Some per-type formatting (spacing, tick marks or markers) could not be applied."
  );
  if (hasExtraFinishing(finishing)) {
    const band = chartBandPosition(baseLeft, baseTop, showY, showX, false);
    await guardedSync(
      ctx,
      () =>
        queueExtraFinishing(chart, finishing, seriesCount, band.width, band.height, defaults.legend),
      warnings,
      finishing.dataLabelFont
        ? undefined // chartex types often refuse label fonts; the VBA skipped this silently too
        : "Some per-type formatting (plot size or doughnut hole) could not be applied."
    );
  }

  return {
    chartName,
    worksheetId: sheet.id,
    groupName: chromeGroupName(chartName),
    warnings,
  };
}
