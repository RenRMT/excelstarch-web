/**
 * Colour ramps, diverging-tag list, and palette-order maps — ported from `reference/spec.json`
 * (`ramps.*`, `divergingTags`, `paletteOrder`). Each ramp is a 10-step hue ladder, index 0 =
 * lightest (step 1) … index 9 = darkest (step 10).
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "./brand";
import { colorNeutral2 } from "./brand";

/** A ramp is exactly ten steps, lightest (index 0) to darkest (index 9). */
export type Ramp = readonly [Hex, Hex, Hex, Hex, Hex, Hex, Hex, Hex, Hex, Hex];

/** The eight ramp names. */
export type RampName = "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H";

export const ramps: Readonly<Record<RampName, Ramp>> = {
  A: [
    "#E6F1F8",
    "#CCE4F1",
    "#99C9E4",
    "#66ADD6",
    "#3392C9",
    "#0077BB",
    "#005F96",
    "#004770",
    "#00304B",
    "#001825",
  ],
  B: [
    "#FFF3F0",
    "#FFE7E0",
    "#FFCFC2",
    "#FFB8A3",
    "#FFA085",
    "#FF8866",
    "#CC6D52",
    "#99523D",
    "#663629",
    "#331B14",
  ],
  C: [
    "#F1FAFF",
    "#E4F5FF",
    "#C9EBFF",
    "#ADE0FF",
    "#92D6FF",
    "#77CCFF",
    "#5FA3CC",
    "#477A99",
    "#305266",
    "#182933",
  ],
  D: [
    "#E6F5F3",
    "#CCEBE7",
    "#99D6CF",
    "#66C2B8",
    "#33ADA0",
    "#009988",
    "#007A6D",
    "#005C52",
    "#003D36",
    "#001F1B",
  ],
  E: [
    "#FFFCEB",
    "#FFF8D6",
    "#FFF1AD",
    "#FFEB85",
    "#FFE45C",
    "#FFDD33",
    "#CCB129",
    "#99851F",
    "#665814",
    "#332C0A",
  ],
  F: [
    "#F7ECE6",
    "#EEDACC",
    "#DDB499",
    "#CC8F66",
    "#BB6933",
    "#AA4400",
    "#883600",
    "#662900",
    "#441B00",
    "#220E00",
  ],
  G: [
    "#F7F5FD",
    "#EEEBFC",
    "#DDD6F8",
    "#CCC2F5",
    "#BBADF1",
    "#AA99EE",
    "#887ABE",
    "#665C8F",
    "#443D5F",
    "#221F30",
  ],
  H: [
    "#F8F8F8",
    "#F1F1F1",
    "#E4E4E4",
    "#D6D6D6",
    "#C9C9C9",
    "#BBBBBB",
    "#969696",
    "#707070",
    "#4B4B4B",
    "#252525",
  ],
};

/** All 42 valid diverging "LEFT|RIGHT" tag combinations (8 ramps, no same-side pairs). */
export const divergingTags: readonly string[] = [
  "A|B",
  "A|C",
  "A|D",
  "A|E",
  "A|F",
  "A|G",
  "B|A",
  "B|C",
  "B|D",
  "B|E",
  "B|F",
  "B|G",
  "C|A",
  "C|B",
  "C|D",
  "C|E",
  "C|F",
  "C|G",
  "D|A",
  "D|B",
  "D|C",
  "D|E",
  "D|F",
  "D|G",
  "E|A",
  "E|B",
  "E|C",
  "E|D",
  "E|F",
  "E|G",
  "F|A",
  "F|B",
  "F|C",
  "F|D",
  "F|E",
  "F|G",
  "G|A",
  "G|B",
  "G|C",
  "G|D",
  "G|E",
  "G|F",
];

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
