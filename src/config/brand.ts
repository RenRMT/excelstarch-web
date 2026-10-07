/**
 * Brand palette — mapped onto the ROOS (RVO Open Ontwerp Systeem) colour tokens in
 * `reference/roos-kleurenkiezer.html`. The data colours are the ROOS series palette verbatim; every
 * other colour is the nearest ROOS token (CIEDE2000) to the value the VBA port originally used, with
 * neutrals kept on the ROOS grey (slate) scale. Values are #RRGGBB strings.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

/** A `#RRGGBB` color string. */
export type Hex = string;

// --- Brand colors --------------------------------------------------------------------------
/** Lintblauw — chart title colour. */
export const colorBrand1: Hex = "#154273";
/** Grijs-900 — subtitle colour. */
export const colorBrand2: Hex = "#0F172A";
/** Zwart; used for dark label/axis text and as the light-fill contrast colour. */
export const colorBrand3: Hex = "#000000";
/** Grijs-050. */
export const colorBrandLightGrey: Hex = "#F8FAFC";
/** Diverging-ramp neutral centre (Grijs-050) — the grey middle series for odd counts. */
export const colorBrand4: Hex = "#F8FAFC";

// --- Neutral colors (ROOS grey scale) ------------------------------------------------------
/** Grijs-200. */
export const colorNeutral1: Hex = "#E2E8F0";
/** Grijs-300. */
export const colorNeutral2: Hex = "#CBD5E1";
/** Grijs-400. */
export const colorNeutral3: Hex = "#94A3B8";
/** Wit. */
export const colorNeutral4: Hex = "#FFFFFF";

/** Semantic alias — use for axis/border white styling and as the dark-fill contrast colour. */
export const colorWhite: Hex = colorNeutral4;

// --- Data series colors (the ROOS reekspalet, in order) ------------------------------------
export const colorData1: Hex = "#007BC7"; // Hemelblauw
export const colorData2: Hex = "#E89440"; // Oranje-750
export const colorData3: Hex = "#CB66A0"; // Violet-600
export const colorData4: Hex = "#6BA549"; // Groen-750
export const colorData5: Hex = "#FFDE94"; // Donkergeel-450
export const colorData6: Hex = "#E06056"; // Rood-750
export const colorData7: Hex = "#8FCAE7"; // Lichtblauw
export const colorData8: Hex = "#94A3B8"; // Grijs-400: always "other" / "previous period"

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
