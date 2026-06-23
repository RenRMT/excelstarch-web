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
 *   Sync E — group the chrome with the chart (must follow the shapes-exist sync).
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { ChartKind } from "../logic/chartType";
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { createChart, retypeChart, positionChartIntoBand } from "./chartFactory";
import { applyChartStyle } from "./chartStyle";
import { colorSeriesByPalette } from "./seriesColorer";
import { buildChrome } from "../chrome/chromeBuilder";
import { groupChromeWithChart, groupChromeOnly, chromeGroupName } from "../chrome/chromeGroup";
import { chromeShapeName } from "../chrome/chromeNames";
import { barChartDefaults, columnChartDefaults } from "../config/chartDefaults";

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
    const defaults = kind === "bar" ? barChartDefaults() : columnChartDefaults();

    // Decide create vs restyle: is a chart currently selected?
    const active = ctx.workbook.getActiveChartOrNullObject();
    active.load("isNullObject");
    await ctx.sync(); // Sync A — know whether to restyle the active chart or create a new one.

    let chart: Excel.Chart;
    if (active.isNullObject) {
      chart = createChart(ctx, ctx.workbook.worksheets.getActiveWorksheet(), kind);
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

    // Resolve canvas origin (reuse an existing canvas's position so a re-run snaps chrome back) and
    // whether a prior chrome group exists (so removeExistingChrome can safely ungroup it).
    const existingCanvas = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, "canvas"));
    existingCanvas.load("left, top, isNullObject");
    const existingGroup = sheet.shapes.getItemOrNullObject(chromeGroupName(chartName));
    existingGroup.load("isNullObject");
    await ctx.sync(); // Sync C — canvas-or-null + group-or-null resolved.

    const baseLeft = existingCanvas.isNullObject ? chart.left : existingCanvas.left;
    const baseTop = existingCanvas.isNullObject ? chart.top : existingCanvas.top;
    const groupExists = !existingGroup.isNullObject;

    // Queue all mutations (no sync between them).
    positionChartIntoBand(
      chart,
      baseLeft,
      baseTop,
      /*showY*/ true,
      /*showX*/ true,
      /*hasLegend*/ false
    );
    applyChartStyle(chart, defaults);
    colorSeriesByPalette(chart, seriesCount);
    const { createdShapeNames, warnings } = buildChrome(
      sheet,
      chartName,
      baseLeft,
      baseTop,
      defaults,
      groupExists
    );
    await ctx.sync(); // Sync D — chart repositioned/styled/coloured AND chrome shapes exist.

    // Group the chrome with the chart. If grouping-with-the-chart is unsupported on this host (or
    // the chart isn't addressable by name in sheet.shapes), fall back to grouping the chrome only.
    // Both attempts are guarded so a grouping failure degrades to a warning rather than discarding
    // an otherwise-complete chart + chrome.
    try {
      groupChromeWithChart(sheet, chartName, createdShapeNames);
      await ctx.sync(); // Sync E — group + rename.
    } catch {
      try {
        groupChromeOnly(sheet, chartName, createdShapeNames);
        await ctx.sync();
        warnings.push(
          "Couldn't group the chart with its chrome on this host; the chrome is grouped but the " +
            "chart moves separately."
        );
      } catch {
        warnings.push(
          "The chart and its chrome were created but could not be grouped on this host."
        );
      }
    }

    return { chartName, groupName: chromeGroupName(chartName), warnings };
  });
}
