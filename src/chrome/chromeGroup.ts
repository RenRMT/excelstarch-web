/* global Excel */
/**
 * Chrome group lifecycle — ported from `GroupChartExChrome` / `RemoveExistingChartExChrome` in
 * `modEngineExChrome.bas`. Removes any prior chrome before a rebuild (so re-runs don't duplicate),
 * and groups the chrome shapes with the chart into one named group.
 *
 * These QUEUE deletes/ungroup/group calls; the caller controls `context.sync()` boundaries.
 *
 * INTEROP: Office.js shape API.
 */
import { chromeSuffix, chromeGroupName } from "./chromeNames";

export { chromeGroupName };

const ALL_SUFFIXES = Object.values(chromeSuffix);

/**
 * Ungroup a prior chrome group (the chart survives the ungroup) and delete every prefixed chrome
 * member by name. Queues mutations only.
 *
 * `groupExists` MUST come from a synced `isNullObject` read: navigating `group.group.ungroup()` on
 * a null-object proxy (no prior group, e.g. a first run) throws when the batch flushes, which would
 * abort the whole create. `delete()` on a null-object member proxy is a documented safe no-op, so
 * the member deletes need no such guard.
 */
export function removeExistingChrome(
  sheet: Excel.Worksheet,
  chartName: string,
  groupExists: boolean
): void {
  if (groupExists) {
    sheet.shapes.getItemOrNullObject(chromeGroupName(chartName)).group.ungroup();
  }

  for (const suffix of ALL_SUFFIXES) {
    sheet.shapes.getItemOrNullObject(chartName + suffix).delete();
  }
}

/**
 * Group the chart with its chrome members into one named group. `memberShapeNames` are the chrome
 * shapes that were actually created (logo / y-axis title may be absent). Returns the group shape.
 * The chart participates by its chart name (Office.js exposes the chart in `sheet.shapes`).
 *
 * Caller wraps this in try/catch: if grouping-with-the-chart is unsupported on the host, it retries
 * with chrome-only and records a warning (chrome still grouped; chart drags separately).
 */
export function groupChromeWithChart(
  sheet: Excel.Worksheet,
  chartName: string,
  memberShapeNames: string[]
): Excel.Shape {
  const group = sheet.shapes.addGroup([chartName, ...memberShapeNames]);
  group.name = chromeGroupName(chartName);
  return group;
}

/** Fallback grouping without the chart (chrome shapes only). */
export function groupChromeOnly(
  sheet: Excel.Worksheet,
  chartName: string,
  memberShapeNames: string[]
): Excel.Shape {
  const group = sheet.shapes.addGroup(memberShapeNames);
  group.name = chromeGroupName(chartName);
  return group;
}
