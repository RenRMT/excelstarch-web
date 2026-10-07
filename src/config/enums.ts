/**
 * Engine enums and styling thresholds (originally ported from the VBA `modConfigDerived.bas`).
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

/** Default formatting for new/reformatted charts. */
export const defaultGridlines = Axis.None;
export const defaultAxisDisplay = Axis.None;
export const defaultLegend = false;
