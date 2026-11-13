/**
 * Pure maths for the composite export: where each layer goes in the output image, the file name,
 * and a word-wrap helper for the fallback text renderer.
 *
 * Worksheet geometry is in points; the output image is in pixels. `scale` multiplies the standard
 * screen resolution (96 px per 72 pt), so scale 2 gives a 1600 px image for a 600 pt canvas.
 *
 * PURE: this module must never import `Excel`/`Office` or touch the DOM.
 */

/** A rectangle in worksheet points, or in pixels once laid out. */
export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Pixels per point at scale 1 (CSS reference pixel: 96 per inch; a point is 1/72 inch). */
export const PX_PER_POINT = 96 / 72;

export const EXPORT_SCALES = [1, 2, 3] as const;
export type ExportScale = (typeof EXPORT_SCALES)[number];

export type ExportFormat = "png" | "jpeg";

/** Pixel placement of each layer inside the output image. */
export interface ExportLayout {
  /** Output image size in whole pixels. */
  width: number;
  height: number;
  /** Pixels per point used (PX_PER_POINT × scale). */
  pxPerPoint: number;
  /** Where the chart image is drawn. */
  chart: Rect;
  /** Where the chrome-group image is drawn, or null when the chart has no chrome. */
  chrome: Rect | null;
  /** Where each annotation image is drawn, in the order given (drawn last, on top). */
  annotations: Rect[];
}

function toPixels(rect: Rect, frame: Rect, k: number): Rect {
  return {
    left: (rect.left - frame.left) * k,
    top: (rect.top - frame.top) * k,
    width: rect.width * k,
    height: rect.height * k,
  };
}

/**
 * Lay the export out. `frame` is the area to capture in worksheet points — the chrome canvas for a
 * styled chart, or the chart itself when it has no chrome. `chart` and `chrome` are the worksheet
 * rectangles of the chart object and the chrome group; both are placed relative to the frame, so
 * anything hanging outside it is clipped.
 */
export function exportLayout(
  frame: Rect,
  chart: Rect,
  chrome: Rect | null,
  scale: number,
  annotations: readonly Rect[] = []
): ExportLayout {
  const k = PX_PER_POINT * scale;
  return {
    width: Math.round(frame.width * k),
    height: Math.round(frame.height * k),
    pxPerPoint: k,
    chart: toPixels(chart, frame, k),
    chrome: chrome ? toPixels(chrome, frame, k) : null,
    annotations: annotations.map((a) => toPixels(a, frame, k)),
  };
}

/** Characters Windows and macOS refuse in file names, plus control characters. */
const ILLEGAL_FILENAME_CHARS = /[<>:"/\\|?*\u0000-\u001f]/g;
const MAX_FILENAME_LENGTH = 80;

/**
 * A safe file name for the export: the chart's title when it has a real one, otherwise the chart's
 * name. `title` is the title box text (`null` when there is no box); a title still equal to
 * `placeholder`, or empty, is not used.
 */
export function exportFileName(
  title: string | null,
  placeholder: string,
  chartName: string,
  format: ExportFormat
): string {
  const useTitle = title !== null && title.trim().length > 0 && title !== placeholder;
  const base = (useTitle ? (title as string) : chartName)
    .replace(ILLEGAL_FILENAME_CHARS, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/, "") // Windows drops trailing dots/spaces
    .slice(0, MAX_FILENAME_LENGTH)
    .trim();
  const extension = format === "jpeg" ? "jpg" : "png";
  return `${base.length > 0 ? base : "chart"}.${extension}`;
}

/**
 * Break `text` into lines no wider than `maxWidth`, measuring with `measure` (e.g. a canvas
 * `measureText`). Explicit line breaks are kept; a single word wider than the box gets a line of
 * its own rather than being split.
 */
export function wrapText(
  text: string,
  maxWidth: number,
  measure: (s: string) => number
): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split(/\r\n|[\r\n\v]/)) {
    let line = "";
    for (const word of paragraph.split(" ")) {
      const candidate = line.length > 0 ? `${line} ${word}` : word;
      if (line.length > 0 && measure(candidate) > maxWidth) {
        lines.push(line);
        line = word;
      } else {
        line = candidate;
      }
    }
    lines.push(line);
  }
  return lines;
}
