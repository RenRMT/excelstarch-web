/**
 * Chart font styling (originally ported from the VBA `modConfig.bas`).
 * Sizes are in points; colours are resolved from the brand palette (not re-hardcoded).
 *
 * Cross-platform note: italic is applied via `font.italic = true` at the interop layer, NEVER a
 * named italic family ("Verdana Italic") — named italic families do not resolve on Mac.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "./brand";
import { colorBrand1, colorBrand2, colorBrand3 } from "./brand";

/** The ROOS typeface. */
export const fontPrimary = "Verdana";

// --- Sizes ---------------------------------------------------------------------------------
export const titleFontSize = 28;
export const subtitleFontSize = 22;
export const axisFontSize = 18;
export const sourceFontSize = 14;

// --- Colours (from the brand palette) ------------------------------------------------------
export const titleFontColor: Hex = colorBrand1;
export const subtitleFontColor: Hex = colorBrand2;
export const axisFontColor: Hex = colorBrand3;
export const sourceFontColor: Hex = colorBrand3;
