/**
 * Brand palette — ported from `reference/spec.json` (`colors.*`), which is the reviewed,
 * BGR→hex-converted snapshot of `modConfig.bas` on the pinned `demo/inso-brand-colors` commit.
 * Values are #RRGGBB strings; the VBA original stored Excel BGR `Long`s, but the whole port
 * works in hex so there is no per-call BGR conversion.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

/** A `#RRGGBB` color string. */
export type Hex = string;

export const orgName = "COMPANY" as const;

// --- Brand colors (spec.colors.brand) -----------------------------------------------------
export const colorBrand1: Hex = "#1B4BA7";
export const colorBrand2: Hex = "#07142C";
/** Near-black; used for dark label/axis text and as the light-fill contrast colour. */
export const colorBrand3: Hex = "#02020A";
export const colorBrandLightGrey: Hex = "#F9F9F9";
/** Diverging-ramp neutral centre (#F9F9F9) — the grey middle series for odd counts. */
export const colorBrand4: Hex = "#F9F9F9";

// --- Neutral colors (spec.colors.neutral) --------------------------------------------------
export const colorNeutral1: Hex = "#DDDDDD";
export const colorNeutral2: Hex = "#BBBBBB";
export const colorNeutral3: Hex = "#9C9C9C";
export const colorNeutral4: Hex = "#FFFFFF";

/** Semantic alias — use for axis/border white styling and as the dark-fill contrast colour. */
export const colorWhite: Hex = colorNeutral4;

// --- Data series colors (spec.colors.data) -------------------------------------------------
export const colorData1: Hex = "#0077BB";
export const colorData2: Hex = "#FF8866";
export const colorData3: Hex = "#77CCFF";
export const colorData4: Hex = "#009988";
export const colorData5: Hex = "#FFDD33";
export const colorData6: Hex = "#AA4400";
export const colorData7: Hex = "#AA99EE";
export const colorData8: Hex = "#BBBBBB";

/** The eight data colours in declaration order, 0-based. */
export const dataColors: readonly Hex[] = [
  colorData1,
  colorData2,
  colorData3,
  colorData4,
  colorData5,
  colorData6,
  colorData7,
  colorData8,
];
