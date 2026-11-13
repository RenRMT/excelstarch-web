/* global Excel */
/**
 * The single end-to-end "create / restyle a branded chart" action — composes the factory, styler,
 * series colourer, and the chrome builder inside ONE `Excel.run`, with explicit sync boundaries.
 * This is the orchestrator the task pane calls.
 *
 * Re-run semantics: if a chart is selected it is retyped + restyled in place (so the chrome snaps
 * back to its existing canvas origin — the reuse-canvas trick) and the text the user typed into the
 * chrome boxes is carried over; otherwise a new chart is created from the selected range.
 *
 * Sync structure (never sync inside a loop):
 *   Sync A — read whether a chart is selected (create vs restyle decision).
 *   Sync B — create/retype the chart; load its name, position, parent worksheet, series count.
 *   Sync C — chrome snapshot (canvas origin, prior group, existing text) + pie/treemap point count.
 *   Sync D — reposition + style + colour the chart AND build all chrome shapes (they now exist).
 *   Sync E — group the chrome shapes (must follow the shapes-exist sync). The live chart can't be
 *            a group member (disjoint Chart/Shape object models); it's sized into the canvas band.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import type { ChartKind } from "../logic/chartType";
import { chartSpecs } from "../config/chartSpecs";
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { createChart, resolveChartRange, retypeChart, positionChartIntoBand } from "./chartFactory";
import { applyChartStyle } from "./chartStyle";
import { colorByPalette } from "./seriesColorer";
import { buildChrome } from "../chrome/chromeBuilder";
import { groupChrome, chromeGroupName } from "../chrome/chromeGroup";
import { readChromeSnapshot } from "../chrome/chromeSnapshot";
import { Axis } from "../config/enums";

export interface CreateChartResult {
  chartName: string;
  /** Id of the worksheet the chart lives on (chart names are only unique per sheet). */
  worksheetId: string;
  groupName: string;
  warnings: string[];
}

export async function createBrandedChart(kind: ChartKind): Promise<RunResult<CreateChartResult>> {
  if (!shapesSupported()) {
    return unsupported<CreateChartResult>();
  }

  return runExcel<CreateChartResult>(async (ctx) => {
    const spec = chartSpecs[kind];
    const defaults = spec.defaults();

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
    // chrome must be built on that sheet.
    const sheet = chart.worksheet;
    sheet.load("id");
    chart.load("name, left, top");
    chart.series.load("count");
    await ctx.sync(); // Sync B — chartName, position, series count, sheet.

    const chartName = chart.name;
    const seriesCount = chart.series.count;

    // Pie/doughnut and treemap have a single series whose slices/tiles are POINTS — colour per
    // point, not per series (a solid series fill would flatten the whole chart to one colour).
    const colourByPoint = spec.colourTarget === "point" && seriesCount > 0;
    const firstSeriesPoints = colourByPoint ? chart.series.getItemAt(0).points : undefined;
    firstSeriesPoints?.load("count");

    // Existing chrome: canvas origin (so a re-run snaps back), prior group (so it can be safely
    // ungrouped) and the text already typed into the boxes (so a re-run keeps it).
    const snapshot = await readChromeSnapshot(ctx, sheet, chartName); // Sync C.

    const baseLeft = snapshot.canvasOrigin ? snapshot.canvasOrigin.left : chart.left;
    const baseTop = snapshot.canvasOrigin ? snapshot.canvasOrigin.top : chart.top;

    // Queue all mutations (no sync between them). Derive the band layout from this kind's defaults
    // so axis-less types (e.g. pie) inset correctly — `axisDisplay` is a bit-flag.
    positionChartIntoBand(
      chart,
      baseLeft,
      baseTop,
      /*showY*/ (defaults.axisDisplay & Axis.Y) !== 0,
      /*showX*/ (defaults.axisDisplay & Axis.X) !== 0
    );
    applyChartStyle(chart, defaults);
    colorByPalette(
      chart,
      spec.colourTarget,
      firstSeriesPoints ? firstSeriesPoints.count : seriesCount
    );
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

    return {
      chartName,
      worksheetId: sheet.id,
      groupName: chromeGroupName(chartName),
      warnings,
    };
  });
}
