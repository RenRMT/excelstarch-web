/**
 * The state machines behind the chart toggles — ported from `modEngineToggles.bas`. Each toggle
 * reads the chart's current state, takes one step through a fixed cycle, and applies the result.
 * The reading and applying are interop (`excel/toggleFlow`); the cycles are here, pure and tested.
 *
 * Axis naming follows the chart's own axes, not screen directions: the VALUE axis carries the
 * numbers and the CATEGORY axis the labels. (The VBA called value-axis gridlines "horizontal", which
 * is only true for column-type charts; on a bar chart they run vertically.)
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

/** Which of the two axes a two-axis setting is on for. */
export interface AxisPair {
  value: boolean;
  category: boolean;
}

/**
 * Gridlines: none → value → category → both → none (the VBA's
 * None → Horizontal → Vertical → Both).
 */
export function nextGridlines(current: AxisPair): AxisPair {
  if (!current.value && !current.category) return { value: true, category: false };
  if (current.value && !current.category) return { value: false, category: true };
  if (!current.value && current.category) return { value: true, category: true };
  return { value: false, category: false };
}

/** Axis tick labels: none → category (X) → value (Y) → both → none. */
export function nextAxisLabels(current: AxisPair): AxisPair {
  if (!current.value && !current.category) return { value: false, category: true };
  if (!current.value && current.category) return { value: true, category: false };
  if (current.value && !current.category) return { value: true, category: true };
  return { value: false, category: false };
}

/**
 * Axis titles: both → Y only → X only → none → both. `value` is the y-axis title box, `category`
 * the x-axis title box. A title for an axis the chart does not have is clamped off, so a one-axis
 * chart never lands on a step where clicking appears to do nothing — except the unavoidable no-op
 * when the chart has neither axis.
 */
export function nextAxisTitles(current: AxisPair, available: AxisPair): AxisPair {
  let next: AxisPair;
  if (current.value && current.category) next = { value: true, category: false };
  else if (current.value) next = { value: false, category: true };
  else if (current.category) next = { value: false, category: false };
  else next = { value: true, category: true };

  const clamped = {
    value: next.value && available.value,
    category: next.category && available.category,
  };
  // If clamping made this step identical to the current state, take one more step.
  if (
    (available.value || available.category) &&
    clamped.value === current.value &&
    clamped.category === current.category
  ) {
    return nextAxisTitles(clamped, available);
  }
  return clamped;
}

export type DataLabelState = "none" | "outside" | "inside" | "other";

/**
 * Data labels: none → outside end → inside centre → none. Stacked charts have no usable outside
 * position, so they cycle none ↔ inside. Labels in any other position count as "other" and are
 * switched off.
 */
export function nextDataLabels(current: DataLabelState, centerOnly: boolean): DataLabelState {
  if (centerOnly) {
    return current === "none" ? "inside" : "none";
  }
  if (current === "none") return "outside";
  if (current === "outside") return "inside";
  return "none";
}

/** Map an Office.js data-label position to the toggle state. */
export function dataLabelStateFrom(hasLabels: boolean, position: string | undefined): DataLabelState {
  if (!hasLabels) return "none";
  if (position === "OutsideEnd") return "outside";
  if (position === "Center") return "inside";
  return "other";
}

const CENTER_ONLY_TYPES: readonly string[] = [
  "ColumnStacked",
  "ColumnStacked100",
  "BarStacked",
  "BarStacked100",
  "AreaStacked",
  "AreaStacked100",
];

/** True for chart types with no usable outside-end label position (stacked types). */
export function isCenterOnlyLabelType(chartType: string): boolean {
  return CENTER_ONLY_TYPES.includes(chartType);
}

/** Short user-facing description of a two-axis state, for the status line. */
export function describeAxisPair(state: AxisPair, valueName: string, categoryName: string): string {
  if (state.value && state.category) return "both";
  if (state.value) return valueName;
  if (state.category) return categoryName;
  return "none";
}
