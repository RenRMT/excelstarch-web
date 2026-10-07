/**
 * Pure series-fill ordering — the colour decisions behind each recolour mode, returned as the full
 * ordered list of `#RRGGBB` fills (one per series). Mirrors the VBA `BuildColorRamp` /
 * `BuildDivergingRamp` / `FormatSeriesColors` loops, but as pure math: the `excel/` layer applies
 * the returned list to the host with a dumb index loop, so the ordering is unit-proven without a
 * host. Composes the already-tested primitives in `colorRamp`/`colorSeries`.
 *
 * Step indices from `colorRamp` are 1-based palette steps (1 = lightest … 6 = darkest); a ramp
 * tuple is indexed `ramp[step - 1]`.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "../config/brand";
import { colorBrand4 } from "../config/brand";
import { ramps, type RampName } from "../config/ramps";
import { getPaletteColor } from "./colorSeries";
import {
  orderedRampSteps,
  priorityStepsSorted,
  divergingSideCount,
  divergingHasMiddle,
  parseDivergingTag,
} from "./colorRamp";

/** Max series a single-hue ramp can colour (the ramp has six steps). */
export const MAX_SINGLE_RAMP_SERIES = 6;
/** Max series a diverging ramp can colour (six steps per side + an optional grey centre). */
export const MAX_DIVERGING_SERIES = 13;

/** Brand-palette fills for series 1..count (Contrasting by default, Rainbow when `useAltOrder`). */
export function paletteFills(count: number, useAltOrder = false): Hex[] {
  const fills: Hex[] = [];
  for (let i = 1; i <= count; i++) {
    fills.push(getPaletteColor(i, useAltOrder));
  }
  return fills;
}

/**
 * Single-hue ramp fills: `count` steps of `rampName`, darkest first (series 1 darkest).
 * Throws `RangeError` when `count > MAX_SINGLE_RAMP_SERIES` (the interop layer surfaces this as the
 * non-blocking "too many series" message).
 */
export function singleRampFills(rampName: RampName, count: number): Hex[] {
  if (count > MAX_SINGLE_RAMP_SERIES) {
    throw new RangeError(
      `A single-hue ramp supports at most ${MAX_SINGLE_RAMP_SERIES} series (got ${count}).`
    );
  }
  const ramp = ramps[rampName];
  return orderedRampSteps(count).map((step) => ramp[step - 1]);
}

/**
 * Diverging-ramp fills for `tag` ("LEFT|RIGHT", e.g. "Rood|Groen"): left ramp dark→light, a grey centre
 * (`colorBrand4`) when the series count is odd, then right ramp light→dark — matching the VBA
 * `BuildDivergingRamp`. The caller must normalise the tag (trim) before calling, like the
 * pure `parseDivergingTag` boundary expects. Throws `RangeError` for an invalid tag or
 * `count > MAX_DIVERGING_SERIES`.
 */
export function divergingFills(tag: string, count: number): Hex[] {
  if (count > MAX_DIVERGING_SERIES) {
    throw new RangeError(
      `A diverging ramp supports at most ${MAX_DIVERGING_SERIES} series (got ${count}).`
    );
  }
  const parsed = parseDivergingTag(tag);
  if (parsed === null) {
    throw new RangeError(`Invalid diverging tag "${tag}" (expected "LEFT|RIGHT", e.g. "Rood|Groen").`);
  }
  const left = ramps[parsed.left as RampName];
  const right = ramps[parsed.right as RampName];
  if (left === undefined || right === undefined) {
    throw new RangeError(`Unknown ramp in diverging tag "${tag}".`);
  }

  const sideCount = divergingSideCount(count);
  const sideSteps = priorityStepsSorted(sideCount); // ascending: 1 lightest .. darkest
  const fills: Hex[] = [];

  // Left side: descending through the sorted steps (dark → light).
  for (let i = sideCount; i >= 1; i--) {
    fills.push(left[sideSteps[i - 1] - 1]);
  }
  // Centre: grey for an odd series count.
  if (divergingHasMiddle(count)) {
    fills.push(colorBrand4);
  }
  // Right side: ascending through the sorted steps (light → dark).
  for (let i = 1; i <= sideCount; i++) {
    fills.push(right[sideSteps[i - 1] - 1]);
  }

  return fills;
}

/** Invert a fill assignment by reversing the current per-series fills. */
export function invertFills(current: readonly Hex[]): Hex[] {
  return current.slice().reverse();
}
