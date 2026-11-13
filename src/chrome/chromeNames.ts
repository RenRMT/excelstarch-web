/**
 * Naming conventions for chrome shapes and their group, shared by the builder, the grouper, and
 * the chart-text writer so they always agree. Shape members are `<chartName><suffix>` (e.g.
 * "Chart 1_TitleBox") so several charts on one sheet do not collide; the group is
 * `ESChromeGroup_<chartName>` so it can be found again on re-run.
 *
 * (The VBA used an `ESTreemapGroup_` prefix for back-compat; this fresh web port standardizes on
 * `ESChromeGroup_`.)
 *
 * PURE: no `Excel`/`Office`.
 */
export const chromeSuffix = {
  canvas: "_Canvas",
  title: "_TitleBox",
  subtitle: "_SubTitleBox",
  source: "_SourceBox",
  yAxis: "_YAxisTitle",
  xAxis: "_XAxisTitle",
} as const;

/**
 * Shapes earlier versions built but no longer do (figure-number box, logo image). Kept only so a
 * rebuild of an older chart deletes them instead of leaving them orphaned.
 */
export const legacyChromeSuffixes: readonly string[] = ["_FigureBox", "_LogoImage"];

export type ChromeSuffixKey = keyof typeof chromeSuffix;

/** The chrome shapes that hold text (everything except the canvas). */
export type ChromeTextKey = Exclude<ChromeSuffixKey, "canvas">;

export const chromeTextKeys: readonly ChromeTextKey[] = [
  "title",
  "subtitle",
  "source",
  "yAxis",
  "xAxis",
];

/**
 * Prefix shared by every chrome member of a chart (`"Chart 1_"`), including ones not listed in
 * `chromeSuffix`, such as annotation boxes. Used to find all members when regrouping.
 */
export function chromeMemberPrefix(chartName: string): string {
  return chartName + "_";
}

/** All chrome member shape names for a chart, in z-order (canvas first). */
export function chromeShapeName(chartName: string, key: ChromeSuffixKey): string {
  return chartName + chromeSuffix[key];
}

const GROUP_PREFIX = "ESChromeGroup_";

export function chromeGroupName(chartName: string): string {
  return GROUP_PREFIX + chartName;
}
