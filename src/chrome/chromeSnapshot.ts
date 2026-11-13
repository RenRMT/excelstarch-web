/* global Excel */
/**
 * Read what chrome a chart already has: the canvas origin, whether the chrome group exists, and
 * the text currently in each text box. A rebuild uses this to snap back to the same canvas
 * position and to carry the user's text over instead of resetting it to the placeholders.
 *
 * Unlike the other chrome modules this one SYNCS (twice): existence first, then the text of the
 * boxes that exist. Navigating into `textFrame` on a null-object shape proxy throws when the batch
 * flushes, so the text can only be loaded once existence is known.
 *
 * INTEROP: Office.js shape API.
 */
import { chromeGroupName, chromeShapeName, chromeTextKeys, ChromeTextKey } from "./chromeNames";

/** The text of each chrome text box that exists (a missing key means the box does not exist). */
export type ChromeTexts = Partial<Record<ChromeTextKey, string>>;

export interface ChromeSnapshot {
  /** Top-left of the existing canvas, or null when the chart has no chrome yet. */
  canvasOrigin: { left: number; top: number } | null;
  groupExists: boolean;
  texts: ChromeTexts;
}

/** Text frames report line breaks as \r, \n or \v depending on host; normalise to \n. */
function normaliseLineBreaks(text: string): string {
  return text.replace(/\r\n|[\r\v]/g, "\n");
}

export async function readChromeSnapshot(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chartName: string
): Promise<ChromeSnapshot> {
  const canvas = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, "canvas"));
  canvas.load("left, top, isNullObject");
  const group = sheet.shapes.getItemOrNullObject(chromeGroupName(chartName));
  group.load("isNullObject");
  const boxes = chromeTextKeys.map((key) => {
    const shape = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, key));
    shape.load("isNullObject");
    return { key, shape };
  });
  await ctx.sync(); // existence of canvas, group and each text box.

  const present = boxes.filter((b) => !b.shape.isNullObject);
  const ranges = present.map((b) => {
    const range = b.shape.textFrame.textRange;
    range.load("text");
    return { key: b.key, range };
  });
  if (ranges.length > 0) {
    await ctx.sync(); // text of the boxes that exist.
  }

  const texts: ChromeTexts = {};
  for (const { key, range } of ranges) {
    texts[key] = normaliseLineBreaks(range.text);
  }

  return {
    canvasOrigin: canvas.isNullObject ? null : { left: canvas.left, top: canvas.top },
    groupExists: !group.isNullObject,
    texts,
  };
}
