/**
 * Colour ramps, diverging-tag list, and palette-order maps. The ramps are the ROOS tint ramps from
 * `reference/roos-palette.json`: each is six defined tokens, index 0 = lightest (tint 150) …
 * index 5 = the base colour at full strength. ROOS defines no tints darker than the base, so the
 * base is the darkest step.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "./brand";
import { colorNeutral2 } from "./brand";

/** A ramp is exactly six steps: tints 150 · 300 · 450 · 600 · 750 · base. */
export type Ramp = readonly [Hex, Hex, Hex, Hex, Hex, Hex];

/** The nine ROOS ramp names; doubles as the display name. */
export type RampName =
  | "Lintblauw"
  | "Violet"
  | "Lichtblauw"
  | "Hemelblauw"
  | "Donkerblauw"
  | "Oranje"
  | "Groen"
  | "Donkergeel"
  | "Rood";

export const ramps: Readonly<Record<RampName, Ramp>> = {
  // Concern
  Lintblauw: ["#DCE3EA", "#B9C7D5", "#95AAC0", "#738EAB", "#507196", "#154273"],
  Violet: ["#F2D9E7", "#E5B2CF", "#D88CB7", "#CB66A0", "#BE4088", "#A90061"],
  Lichtblauw: ["#EEF7FC", "#DEF0F8", "#CCE7F4", "#BCDFF1", "#ABD7ED", "#8FCAE7"],
  // Functioneel
  Hemelblauw: ["#D9EBF7", "#B3D7EE", "#8CC4E6", "#66B0DD", "#409CD5", "#007BC7"],
  Donkerblauw: ["#D9E9F0", "#B3D2E1", "#8CBBD2", "#67A4C3", "#418EB4", "#01689B"],
  // Signaal
  Oranje: ["#FBEAD9", "#F6D4B3", "#F1BF8C", "#EDA966", "#E89440", "#E17000"],
  Groen: ["#E2EDDB", "#C4DBB7", "#A6C991", "#88B76D", "#6BA549", "#39870C"],
  Donkergeel: ["#FFF4DB", "#FFE9B8", "#FFDE94", "#FFD371", "#FFC84D", "#FFB612"],
  Rood: ["#F9DFDD", "#F3C0BC", "#EC9F99", "#E68078", "#E06056", "#D51B1E"],
};

/** The ramp names in menu order (ROOS grouping: concern, functional, signal). */
export const rampOrder: readonly RampName[] = Object.keys(ramps) as RampName[];

/** Every valid diverging "LEFT|RIGHT" tag: each ordered pair of two different ramps. */
export const divergingTags: readonly string[] = rampOrder.reduce<string[]>(
  (tags, left) =>
    tags.concat(rampOrder.filter((right) => right !== left).map((right) => `${left}|${right}`)),
  []
);

/**
 * Palette-slot ordering for series colouring (1-based slot → data-colour index 1..8).
 * Contrasting is the default; Rainbow is the alternate toggle order.
 */
export const paletteOrder = {
  contrasting: [1, 2, 3, 4, 5, 6, 7, 8] as readonly number[],
  rainbow: [1, 7, 3, 4, 5, 2, 6, 8] as readonly number[],
  /** Fallback fill for series beyond the eighth. */
  fallbackOver8: colorNeutral2,
} as const;
