/* global Excel */
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
/**
 * Read the data points of one series of a chart, for the pane's point pickers (point fill,
 * annotation anchor). Point names are the category labels; when those cannot be read, points are
 * numbered.
 *
 * SYNCS. INTEROP: Office.js chart API.
 */

/** Names of the points of `series` (category labels, or "Point n" as a fallback). */
export async function readPointNames(
  ctx: Excel.RequestContext,
  series: Excel.ChartSeries
): Promise<string[]> {
  const points = series.points;
  points.load("count");
  await ctx.sync();
  const fallback = Array.from({ length: points.count }, (_, i) => `Point ${i + 1}`);
  if (points.count === 0) {
    return fallback;
  }
  // Reading categories can fail on some chart types; keep the numbered fallback then.
  try {
    const categories = series.getDimensionValues("Categories");
    await ctx.sync();
    if (categories.value.length === points.count) {
      return categories.value.map((c, i) => (c ? String(c) : fallback[i]));
    }
  } catch {
    // fall through
  }
  return fallback;
}

/** Point names of one series of the selected chart (empty when there is no such series). */
export async function listPointNames(series: number): Promise<RunResult<string[]>> {
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
    chart.series.load("count");
    await ctx.sync();
    if (series < 0 || series >= chart.series.count) {
      return [];
    }
    return readPointNames(ctx, chart.series.getItemAt(series));
  });
}
