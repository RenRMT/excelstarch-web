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
  figure: "_FigureBox",
  title: "_TitleBox",
  subtitle: "_SubTitleBox",
  source: "_SourceBox",
  logo: "_LogoImage",
  yAxis: "_YAxisTitle",
} as const;

export type ChromeSuffixKey = keyof typeof chromeSuffix;

/** All chrome member shape names for a chart, in z-order (canvas first). */
export function chromeShapeName(chartName: string, key: ChromeSuffixKey): string {
  return chartName + chromeSuffix[key];
}

const GROUP_PREFIX = "ESChromeGroup_";

export function chromeGroupName(chartName: string): string {
  return GROUP_PREFIX + chartName;
}
