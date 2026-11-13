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
import {
  chromeSuffix,
  chromeGroupName,
  chromeMemberPrefix,
  legacyChromeSuffixes,
} from "./chromeNames";

export { chromeGroupName };

const ALL_SUFFIXES = [...Object.values(chromeSuffix), ...legacyChromeSuffixes];

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
 * Group the chrome members into one named group. `memberShapeNames` are the chrome shapes that were
 * actually created (the y-axis title may be absent). Returns the group shape.
 *
 * The live chart is intentionally NOT a member: Office.js has disjoint Chart and Shape object models
 * — a chart is not addressable in `sheet.shapes` (no `chart` value in `Excel.ShapeType`) and
 * `Excel.Chart` exposes no group/shape handle, so it cannot join a shape group. The chart is instead
 * sized into the canvas band (`positionChartIntoBand`) so it sits within the chrome visually; users
 * move the pair by selecting the chart together with the group.
 */
export function groupChrome(
  sheet: Excel.Worksheet,
  chartName: string,
  memberShapeNames: string[]
): Excel.Shape {
  const group = sheet.shapes.addGroup(memberShapeNames);
  group.name = chromeGroupName(chartName);
  return group;
}

/**
 * Group every ungrouped shape on the sheet whose name starts with the chart's member prefix — the
 * built chrome plus anything added later (axis titles, annotations). SYNCS twice: once to list the
 * shapes, once to group. Call it after the members exist and after any previous group has been
 * ungrouped (shapes inside a group are not listed at the top level). Grouping needs at least two
 * shapes; with fewer it does nothing.
 */
export async function regroupChrome(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chartName: string
): Promise<void> {
  const shapes = sheet.shapes;
  shapes.load("items/name");
  await ctx.sync();

  const prefix = chromeMemberPrefix(chartName);
  const members = shapes.items.map((s) => s.name).filter((name) => name.startsWith(prefix));
  if (members.length < 2) {
    return;
  }
  groupChrome(sheet, chartName, members);
  await ctx.sync();
}
