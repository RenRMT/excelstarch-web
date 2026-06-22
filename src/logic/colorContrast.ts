/**
 * Pure colour-contrast utilities — ported from `modColorContrast.bas`.
 *
 * Chooses readable data-label text (white vs dark brand) against a coloured fill, following
 * the WCAG 2.x relative-luminance definition. The VBA original took an Excel BGR `Long`; this
 * port takes a `#RRGGBB` hex string and parses the channels itself, so the maths is identical.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "../config/brand";
import { colorWhite, colorBrand3 } from "../config/brand";
import { wcagLuminanceThreshold } from "../config/enums";

/** Parse a `#RRGGBB` (or `RRGGBB`) hex string into 8-bit R/G/B channels. */
function hexToRgb(hex: Hex): { r: number; g: number; b: number } {
  const h = hex.replace(/^#/, "");
  return {
    r: parseInt(h.slice(0, 2), 16),
    g: parseInt(h.slice(2, 4), 16),
    b: parseInt(h.slice(4, 6), 16),
  };
}

/** Linearize one sRGB channel (0..1) to its WCAG linear value. */
function linearize(channel: number): number {
  return channel <= 0.03928 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4);
}

/**
 * WCAG relative luminance of a colour, in [0, 1] (0 = black, 1 = white).
 * Channel weights and the sRGB-linearization constants are the standard WCAG 2.x values.
 */
export function relativeLuminance(hex: Hex): number {
  const { r, g, b } = hexToRgb(hex);
  const rs = linearize(r / 255);
  const gs = linearize(g / 255);
  const bs = linearize(b / 255);
  return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
}

/**
 * White text on dark fills, dark brand text on light fills, using the WCAG luminance threshold.
 */
export function contrastColorForFill(fill: Hex): Hex {
  return relativeLuminance(fill) < wcagLuminanceThreshold ? colorWhite : colorBrand3;
}
