/* global Excel */
/**
 * Add or remove the axis-title boxes of an already-styled chart and move the chart band to match,
 * without touching anything else (colours, text in the other boxes, annotations). Used by the
 * axis-title toggle — the VBA `ApplyAxisTitleLayout`, which deleted/recreated the title boxes and
 * resized the plot area.
 *
 * A box that stays keeps its text; a box that is added gets its placeholder. Because Office.js
 * cannot add a shape to an existing group, the chrome is ungrouped, changed and regrouped.
 *
 * SYNCS (several times). INTEROP: Office.js shape and chart API.
 */
import type { AxisTitles } from "./chromeBuilder";
import { readChromeSnapshot } from "./chromeSnapshot";
import { chromePositions } from "./chromeLayout";
import { chromeGroupName, chromeShapeName } from "./chromeNames";
import { regroupChrome } from "./chromeGroup";
import { addXAxisTitle, addYAxisTitle } from "./chromeShapes";
import { chartBandPosition } from "./chromeLayout";

export const NOT_STYLED =
  "This chart has no ExcelStarch layout yet. Create it with the builder or click Apply style first.";

/** Which axis-title boxes the chart currently has. Null when the chart has no chrome at all. */
export async function readAxisTitles(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chartName: string
): Promise<AxisTitles | null> {
  const snapshot = await readChromeSnapshot(ctx, sheet, chartName);
  if (snapshot.canvasOrigin === null) {
    return null;
  }
  return { y: snapshot.texts.yAxis !== undefined, x: snapshot.texts.xAxis !== undefined };
}

export async function setAxisTitles(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chart: Excel.Chart,
  chartName: string,
  titles: AxisTitles
): Promise<void> {
  const snapshot = await readChromeSnapshot(ctx, sheet, chartName);
  if (snapshot.canvasOrigin === null) {
    throw new Error(NOT_STYLED);
  }
  const { left, top } = snapshot.canvasOrigin;
  const has = { y: snapshot.texts.yAxis !== undefined, x: snapshot.texts.xAxis !== undefined };

  if (snapshot.groupExists) {
    sheet.shapes.getItem(chromeGroupName(chartName)).group.ungroup();
  }

  const pos = chromePositions(left, top, titles.y, titles.x);
  if (has.y && !titles.y) {
    sheet.shapes.getItem(chromeShapeName(chartName, "yAxis")).delete();
  } else if (!has.y && titles.y && pos.yAxis) {
    addYAxisTitle(sheet, chartName, pos.yAxis);
  }
  if (has.x && !titles.x) {
    sheet.shapes.getItem(chromeShapeName(chartName, "xAxis")).delete();
  } else if (!has.x && titles.x && pos.xAxis) {
    addXAxisTitle(sheet, chartName, pos.xAxis);
  }

  const band = chartBandPosition(left, top, titles.y, titles.x, false);
  chart.left = band.left;
  chart.top = band.top;
  chart.width = band.width;
  chart.height = band.height;
  await ctx.sync();

  await regroupChrome(ctx, sheet, chartName);
}
