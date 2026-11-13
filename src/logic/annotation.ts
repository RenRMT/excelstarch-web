/**
 * Pure placement and naming for annotation boxes — ported from `AnnotateButton` in
 * `modEngineStyle.bas`. A box goes at the centre of the plot, or just above and right of a data
 * point; repeated runs stack (each box gets the next free number) rather than replacing.
 *
 * Positions here are in worksheet points: the caller adds the chart's own position to the
 * chart-relative plot and data-label coordinates Excel reports.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Rect } from "./exportLayout";
import {
  annotationBoxWidth,
  annotationBoxHeight,
  annotationOffsetX,
  annotationOffsetY,
} from "../config/geometry";

const ANNOTATION_INFIX = "_Annotation";

/** The name prefix of a chart's annotation boxes, e.g. `"Chart 1_Annotation"`. */
export function annotationPrefix(chartName: string): string {
  return chartName + ANNOTATION_INFIX;
}

/** True for an annotation box of any chart (used to keep annotations out of the chrome group). */
export function isAnnotationName(name: string): boolean {
  return /_Annotation\d+$/.test(name);
}

/** The next free annotation name for a chart, given the names of the shapes on its sheet. */
export function nextAnnotationName(chartName: string, existingNames: readonly string[]): string {
  const prefix = annotationPrefix(chartName);
  let highest = 0;
  for (const name of existingNames) {
    if (name.startsWith(prefix)) {
      const n = Number(name.slice(prefix.length));
      if (Number.isInteger(n) && n > highest) {
        highest = n;
      }
    }
  }
  return `${prefix}${highest + 1}`;
}

/** The box centred on a rectangle (the plot's inner area, or the chart when that is unknown). */
export function annotationBoxAtCentre(area: Rect): Rect {
  return {
    left: area.left + area.width / 2 - annotationBoxWidth / 2,
    top: area.top + area.height / 2 - annotationBoxHeight / 2,
    width: annotationBoxWidth,
    height: annotationBoxHeight,
  };
}

/** The box beside a data point, positioned from the point's data-label top-left corner. */
export function annotationBoxAtPoint(labelLeft: number, labelTop: number): Rect {
  return {
    left: labelLeft + annotationOffsetX,
    top: labelTop + annotationOffsetY,
    width: annotationBoxWidth,
    height: annotationBoxHeight,
  };
}
