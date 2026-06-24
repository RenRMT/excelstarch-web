/* global Excel */
/**
 * The single end-to-end "create / restyle a branded chart" action — composes the factory, styler,
 * series colourer, and the chrome builder inside ONE `Excel.run`, with explicit sync boundaries.
 * This is the orchestrator the task pane calls.
 *
 * Re-run semantics: if a chart is selected it is retyped + restyled in place (so the chrome snaps
 * back to its existing canvas origin — the reuse-canvas trick); otherwise a new chart is created
 * from the selected range.
 *
 * Sync structure (never sync inside a loop):
 *   Sync A — read whether a chart is selected (create vs restyle decision).
 *   Sync B — create/retype the chart; load its name, position, parent worksheet, series count.
 *   Sync C — resolve the canvas origin and prior-group existence on the chart's own worksheet.
 *   Sync D — reposition + style + colour the chart AND build all chrome shapes (they now exist).
 *   Sync E — group the chrome shapes (must follow the shapes-exist sync). The live chart can't be
 *            a group member (disjoint Chart/Shape object models); it's sized into the canvas band.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { ChartKind } from "../logic/chartType";
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { createChart, resolveChartRange, retypeChart, positionChartIntoBand } from "./chartFactory";
import { applyChartStyle } from "./chartStyle";
import { colorSeriesByPalette, colorPointsByPalette } from "./seriesColorer";
import { buildChrome } from "../chrome/chromeBuilder";
import { groupChrome, chromeGroupName } from "../chrome/chromeGroup";
import { chromeShapeName } from "../chrome/chromeNames";
import type { ChartDefaults } from "../config/chartDefaults";
import {
  barChartDefaults,
  columnChartDefaults,
  lineChartDefaults,
  areaChartDefaults,
  scatterChartDefaults,
  pieChartDefaults,
} from "../config/chartDefaults";
import { Axis } from "../config/enums";

/**
 * Per-kind formatting-defaults factory. Exhaustive — a new ChartKind without an entry is a compile
 * error (the keys are `Record<ChartKind, …>`).
 */
const defaultsForKind: Record<ChartKind, () => ChartDefaults> = {
  bar: barChartDefaults,
  column: columnChartDefaults,
  line: lineChartDefaults,
  area: areaChartDefaults,
  scatter: scatterChartDefaults,
  pie: pieChartDefaults,
};

export interface CreateChartResult {
  chartName: string;
  groupName: string;
  warnings: string[];
}

export async function createBrandedChart(kind: ChartKind): Promise<RunResult<CreateChartResult>> {
  if (!shapesSupported()) {
    return unsupported<CreateChartResult>();
  }

  return runExcel<CreateChartResult>(async (ctx) => {
    const defaults = defaultsForKind[kind]();

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
    // Resolve the chart's OWN worksheet (a selected chart may live on a non-active sheet); the
    // chrome must be built on that sheet so the group can include the chart.
    const sheet = chart.worksheet;
    chart.load("name, left, top");
    chart.series.load("count");
    await ctx.sync(); // Sync B — chartName, position, series count, sheet.

    const chartName = chart.name;
    const seriesCount = chart.series.count;

    // Pie/doughnut have a single series whose slices are POINTS — colour per point, not per series
    // (a solid series fill would flatten the whole pie to one colour). Load that point count here.
    const colourByPoint = kind === "pie";
    const firstSeriesPoints = colourByPoint ? chart.series.getItemAt(0).points : undefined;
    firstSeriesPoints?.load("count");

    // Resolve canvas origin (reuse an existing canvas's position so a re-run snaps chrome back) and
    // whether a prior chrome group exists (so removeExistingChrome can safely ungroup it).
    const existingCanvas = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, "canvas"));
    existingCanvas.load("left, top, isNullObject");
    const existingGroup = sheet.shapes.getItemOrNullObject(chromeGroupName(chartName));
    existingGroup.load("isNullObject");
    await ctx.sync(); // Sync C — canvas-or-null + group-or-null resolved (+ pie point count).

    const baseLeft = existingCanvas.isNullObject ? chart.left : existingCanvas.left;
    const baseTop = existingCanvas.isNullObject ? chart.top : existingCanvas.top;
    const groupExists = !existingGroup.isNullObject;

    // Queue all mutations (no sync between them). Derive the band layout from this kind's defaults
    // so axis-less, legend-on types (e.g. pie) inset correctly — `axisDisplay` is a bit-flag.
    positionChartIntoBand(
      chart,
      baseLeft,
      baseTop,
      /*showY*/ (defaults.axisDisplay & Axis.Y) !== 0,
      /*showX*/ (defaults.axisDisplay & Axis.X) !== 0,
      /*hasLegend*/ defaults.legend
    );
    applyChartStyle(chart, defaults);
    if (firstSeriesPoints) {
      colorPointsByPalette(chart, firstSeriesPoints.count);
    } else {
      colorSeriesByPalette(chart, seriesCount);
    }
    const { createdShapeNames, warnings } = buildChrome(
      sheet,
      chartName,
      baseLeft,
      baseTop,
      defaults,
      groupExists
    );
    await ctx.sync(); // Sync D — chart repositioned/styled/coloured AND chrome shapes exist.

    // Group the chrome shapes. The live chart can't join the group — Office.js has disjoint Chart
    // and Shape object models (see chromeGroup.groupChrome) — so it's sized into the canvas band
    // above and sits within the chrome visually. Guard the group call so a rare failure degrades to
    // a warning rather than discarding an otherwise-complete chart + chrome.
    try {
      groupChrome(sheet, chartName, createdShapeNames);
      await ctx.sync(); // Sync E — group + rename.
    } catch {
      warnings.push("The chart and its chrome were created but the chrome could not be grouped.");
    }

    return { chartName, groupName: chromeGroupName(chartName), warnings };
  });
}
