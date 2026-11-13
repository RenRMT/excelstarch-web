/**
 * Optional logo in the bottom-right corner of the chart canvas — the VBA `InsertLogo`
 * (`modEmbeddedImages.bas`). Off by default: the ROOS house style has no in-chart logo.
 *
 * To turn it on, set `logoEnabled` to true and paste a base64-encoded PNG or JPEG (no `data:`
 * prefix) into `logoBase64`. Office.js can only insert PNG and JPEG pictures, so convert an SVG
 * logo first.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import { chartWidth, chartHeight, sourceBoxWidth } from "./geometry";

export const logoEnabled = false;

/** Base64 PNG/JPEG data of the logo, without a `data:` prefix. */
export const logoBase64 = "";

/** Logo height as a share of the canvas height. */
const logoHeightScale = 0.1;
/** Logo width ÷ height. Set it explicitly: Excel does not keep the aspect ratio reliably. */
const logoAspectRatio = 1;
/** Gap to the canvas's right and bottom edges, as a share of the canvas width/height. */
const logoMarginProportion = 0.01;

export const logoHeight = chartHeight * logoHeightScale;
export const logoWidth = logoHeight * logoAspectRatio;

/**
 * The logo's canvas-relative box: bottom-right, inside the margins. It shares the bottom band with
 * the source box and clears it only because the source box ends left of the logo column.
 */
export function logoBox(): { left: number; top: number; width: number; height: number } {
  return {
    left: chartWidth - logoWidth - chartWidth * logoMarginProportion,
    top: chartHeight - logoHeight - chartHeight * logoMarginProportion,
    width: logoWidth,
    height: logoHeight,
  };
}

/** Whether the logo fits beside the source box (the layout assumption above). */
export function logoClearsSourceBox(): boolean {
  return logoBox().left >= sourceBoxWidth;
}
