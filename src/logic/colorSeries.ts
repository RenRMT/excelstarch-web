/**
 * Pure palette-order lookup — ported from `GetPaletteColor` in `modColorSeries.bas`.
 * The object-model series-colouring loops (`FormatSeriesColors`, `TogglePaletteOrder`) are NOT
 * here — they belong to the Phase-1 `excel/` layer.
 *
 * The VBA original read the palette order from module-level state (`m_useAltOrder`, flipped only
 * by the object-model toggle). Per the testing rule, this port injects the order as a parameter
 * so the alternate branch is unit-testable.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "../config/brand";
import { dataColors } from "../config/brand";
import { paletteOrder } from "../config/ramps";

/**
 * The brand colour for 1-based series index `i`, respecting the palette order.
 * Contrasting (default) maps slot i → dataColor i; Rainbow reorders. Indices outside 1..8 fall
 * back to the neutral (`colorNeutral2`).
 */
export function getPaletteColor(i: number, useAltOrder = false): Hex {
  if (i < 1 || i > 8) return paletteOrder.fallbackOver8;
  const order = useAltOrder ? paletteOrder.rainbow : paletteOrder.contrasting;
  // order[i-1] is a 1-based data-colour index; dataColors is 0-based.
  return dataColors[order[i - 1] - 1];
}
