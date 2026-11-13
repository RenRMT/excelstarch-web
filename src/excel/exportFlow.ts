/* global Excel */
/**
 * Fetch what the composite export needs from the host: the selected chart as an image, its chrome
 * group as an image, and the worksheet geometry of both. The VBA exported with `Chart.Export`;
 * Office.js can only return images as base64, and a chart's image never includes the chrome (the
 * chrome is worksheet shapes, not part of the chart), so the task pane composes the two layers
 * itself — see `taskpane/export/compose`.
 *
 * Also inserts a finished image back onto the worksheet as a picture, the one delivery route that
 * needs no browser download or clipboard support.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { readChromeSnapshot, ChromeTexts } from "../chrome/chromeSnapshot";
import { chromeGroupName } from "../chrome/chromeGroup";
import { chartWidth, chartHeight } from "../config/geometry";
import { Rect, PX_PER_POINT } from "../logic/exportLayout";

export interface ExportLayers {
  chartName: string;
  worksheetId: string;
  /** The area to capture, in worksheet points: the chrome canvas, or the chart if it has none. */
  frame: Rect;
  chart: Rect;
  /** Base64 PNG of the chart object, rendered at the requested scale. */
  chartImage: string;
  /** Worksheet rectangle of the chrome group, or null when the chart has no chrome. */
  chrome: Rect | null;
  /**
   * Base64 PNG of the chrome group as Excel draws it, or null when the chart has no chrome or the
   * host could not render the group (the composer then draws the text itself from `texts`).
   */
  chromeImage: string | null;
  /** Text of the chrome boxes, for the file name and the fallback text renderer. */
  texts: ChromeTexts;
}

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";

/** Read the selected chart and its chrome as images, with their geometry, at `scale`. */
export async function readExportLayers(scale: number): Promise<RunResult<ExportLayers>> {
  if (!shapesSupported()) {
    return unsupported<ExportLayers>();
  }

  return runExcel<ExportLayers>(async (ctx) => {
    const chart = ctx.workbook.getActiveChartOrNullObject();
    chart.load("isNullObject, name, left, top, width, height");
    await ctx.sync(); // active chart + its geometry.
    if (chart.isNullObject) {
      throw new Error(NO_CHART_SELECTED);
    }
    const sheet = chart.worksheet;
    sheet.load("id");

    const chartRect: Rect = {
      left: chart.left,
      top: chart.top,
      width: chart.width,
      height: chart.height,
    };
    const snapshot = await readChromeSnapshot(ctx, sheet, chart.name); // also resolves sheet.id.

    const hasChrome = snapshot.canvasOrigin !== null;
    const frame: Rect = snapshot.canvasOrigin
      ? { ...snapshot.canvasOrigin, width: chartWidth, height: chartHeight }
      : chartRect;

    // Ask for the chart at the output resolution so it is not upscaled when composed.
    const k = PX_PER_POINT * scale;
    const chartImage = chart.getImage(
      Math.round(chartRect.width * k),
      Math.round(chartRect.height * k),
      Excel.ImageFittingMode.fit
    );
    const group = snapshot.groupExists
      ? sheet.shapes.getItem(chromeGroupName(chart.name))
      : undefined;
    group?.load("left, top, width, height");
    await ctx.sync(); // chart image + group geometry.

    // Rendering a group as an image is the least certain call here, so it gets its own guarded
    // sync: if the host refuses, the export still works with pane-drawn text.
    let chromeImage: string | null = null;
    if (group) {
      try {
        const image = group.getAsImage(Excel.PictureFormat.png);
        await ctx.sync();
        chromeImage = image.value;
      } catch {
        chromeImage = null;
      }
    }

    return {
      chartName: chart.name,
      worksheetId: sheet.id,
      frame,
      chart: chartRect,
      chartImage: chartImage.value,
      chrome: group
        ? { left: group.left, top: group.top, width: group.width, height: group.height }
        : hasChrome
          ? frame
          : null,
      chromeImage,
      texts: snapshot.texts,
    };
  });
}

/** Gap in points between the chart's frame and a picture inserted beside it. */
const INSERT_GAP = 20;

/**
 * Insert a finished export onto the chart's worksheet as a picture, to the right of the chart's
 * frame and at the same size. `base64` is the bare image data (no `data:` prefix), PNG or JPEG.
 */
export async function insertExportPicture(
  worksheetId: string,
  frame: Rect,
  base64: string
): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  return runExcel<void>(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(worksheetId);
    const picture = sheet.shapes.addImage(base64);
    picture.lockAspectRatio = false;
    picture.left = frame.left + frame.width + INSERT_GAP;
    picture.top = frame.top;
    picture.width = frame.width;
    picture.height = frame.height;
    picture.lockAspectRatio = true;
    await ctx.sync();
  });
}
