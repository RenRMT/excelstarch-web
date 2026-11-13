/* global Excel */
/**
 * Chrome build orchestrator — ported from `BuildChartExChrome` in `modEngineExChrome.bas`,
 * generalized from chartex to every chart type. Queues: remove prior chrome → white canvas
 * (sent to back) → title/subtitle/source text → optional y-axis title. `texts` carries the text of
 * the boxes being replaced (from `readChromeSnapshot`), so a rebuild keeps what the user typed; a
 * box with no previous text gets its placeholder. Returns the names of the
 * shapes actually created (for the flow to group after the shapes-exist sync) plus any warnings.
 *
 * QUEUES shape creation only — does NOT sync and does NOT group (the flow groups post-sync, once
 * the shapes exist on the host).
 *
 * INTEROP: Office.js shape API.
 */
import type { ChartDefaults } from "../config/chartDefaults";
import { chromePositions } from "./chromeLayout";
import { removeExistingChrome } from "./chromeGroup";
import type { ChromeTexts } from "./chromeSnapshot";
import {
  addCanvas,
  addTitleBox,
  addSubtitleBox,
  addSourceBox,
  addYAxisTitle,
} from "./chromeShapes";

export interface ChromeBuildResult {
  /** Names of the chrome shapes created this run, in z-order (canvas first). */
  createdShapeNames: string[];
  /** Non-fatal issues; currently none are raised, kept for the flow's warning plumbing. */
  warnings: string[];
}

export function buildChrome(
  sheet: Excel.Worksheet,
  chartName: string,
  baseLeft: number,
  baseTop: number,
  defaults: ChartDefaults,
  groupExists: boolean,
  texts: ChromeTexts = {}
): ChromeBuildResult {
  removeExistingChrome(sheet, chartName, groupExists);

  const pos = chromePositions(baseLeft, baseTop, defaults.showYAxisTitle);
  const createdShapeNames: string[] = [];
  const warnings: string[] = [];

  // Canvas first so it sits at the back of the z-order.
  createdShapeNames.push(addCanvas(sheet, chartName, pos.canvas).name);
  createdShapeNames.push(addTitleBox(sheet, chartName, pos.title, texts.title).name);
  createdShapeNames.push(addSubtitleBox(sheet, chartName, pos.subtitle, texts.subtitle).name);
  createdShapeNames.push(addSourceBox(sheet, chartName, pos.source, texts.source).name);

  if (defaults.showYAxisTitle && pos.yAxis) {
    createdShapeNames.push(addYAxisTitle(sheet, chartName, pos.yAxis, texts.yAxis).name);
  }

  return { createdShapeNames, warnings };
}
