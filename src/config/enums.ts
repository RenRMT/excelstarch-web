/**
 * Engine enums and styling thresholds — ported from `reference/spec.json` (`enums.*`),
 * mirroring `modConfigDerived.bas`.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

/** Axis selection: which axes a flag applies to. */
export enum Axis {
  None = 0,
  X = 1,
  Y = 2,
  Both = 3,
}

/** WCAG relative-luminance threshold for choosing black vs white label text. */
export const wcagLuminanceThreshold = 0.179;

/** Marker point size (points) for scatter charts. */
export const scatterMarkerSize = 9;

/** Bubble fill transparency (0 = opaque, 1 = fully clear). */
export const bubbleTransparency = 0.5;

/** Default formatting for new/reformatted charts (spec.enums.defaults). */
export const defaultGridlines = Axis.None;
export const defaultAxisDisplay = Axis.None;
export const defaultLegend = false;
