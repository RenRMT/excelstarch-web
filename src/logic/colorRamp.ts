/**
 * Pure ramp step-ordering logic — ported from the unit-testable functions in `modColorRamp.bas`
 * (`OrderedRampSteps`, `DivergingSideCount`/`HasMiddle`, `ParseDivergingTag`, and the private
 * `StepPriority`/`PriorityStepsSorted`). The object-model fill loops (`BuildColorRamp`,
 * `BuildDivergingRamp`) are NOT here — they belong to the Phase-1 `excel/` layer.
 *
 * Step indices are 1-based palette steps (1 = lightest … 6 = darkest, the ROOS base colour),
 * matching the ramp tuples when indexed as `ramp[step - 1]`.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

/**
 * The fixed step-selection priority: which palette steps to use, and in what preference order,
 * as the series count grows. Shared by single and diverging ramps. The base colour (6) comes first;
 * the lightest tint (1) last.
 */
const STEP_PRIORITY: readonly number[] = [6, 2, 4, 3, 5, 1];

/**
 * The first `count` priority steps, sorted ascending (1 = lightest … 6 = darkest).
 * `count <= 0` yields an empty array. `count` must be 0..6.
 */
export function priorityStepsSorted(count: number): number[] {
  if (count <= 0) return [];
  return STEP_PRIORITY.slice(0, count).sort((a, b) => a - b);
}

/**
 * Single-hue ramp: the 1-based palette step for each series, darkest first (series 1 gets the
 * darkest selected step, series n the lightest). Returns an n-element array; `n <= 0` yields `[]`.
 */
export function orderedRampSteps(n: number): number[] {
  if (n <= 0) return [];
  // Ascending (lightest..darkest), then reversed so the darkest comes first.
  return priorityStepsSorted(n).reverse();
}

/** Diverging ramp: number of series on each side (floor(n / 2)). */
export function divergingSideCount(n: number): number {
  return Math.floor(n / 2);
}

/** Diverging ramp: true when n is odd, so a grey centre series is inserted. */
export function divergingHasMiddle(n: number): boolean {
  return n % 2 === 1;
}

/**
 * Parse a diverging-ramp tag `"LEFT|RIGHT"` (e.g. `"A|B"`) into its two ramp names.
 * Returns `null` when the pipe separator or either side is missing (the VBA original signalled
 * this with a `Boolean` return + `ByRef` out-params; we return a typed result or `null` instead).
 *
 * This is the pure boundary and does NOT upper-case or trim the input — the VBA caller applied
 * `UCase$`/`Trim$` before calling. The Phase-1 interop caller must normalize the tag the same way.
 */
export function parseDivergingTag(tag: string): { left: string; right: string } | null {
  const parts = tag.split("|");
  if (parts.length < 2) return null; // no separator → invalid
  if (parts[0].length === 0 || parts[1].length === 0) return null;
  return { left: parts[0], right: parts[1] };
}
