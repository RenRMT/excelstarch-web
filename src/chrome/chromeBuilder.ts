/* global Excel */
/**
 * Chrome build orchestrator — ported from `BuildChartExChrome` in `modEngineExChrome.bas`,
 * generalized from chartex to every chart type. Queues: remove prior chrome → white canvas
 * (sent to back) → figure/title/subtitle/source text → logo → optional y-axis title. Returns the
 * names of the shapes actually created (for the flow to group after the shapes-exist sync) plus
 * any warnings (e.g. the logo failed to add).
 *
 * QUEUES shape creation only — does NOT sync and does NOT group (the flow groups post-sync, once
 * the shapes exist on the host).
 *
 * INTEROP: Office.js shape API.
 */
import type { ChartDefaults } from "../config/chartDefaults";
import { chromePositions } from "./chromeLayout";
import { removeExistingChrome } from "./chromeGroup";
import {
  addCanvas,
  addFigureBox,
  addTitleBox,
  addSubtitleBox,
  addSourceBox,
  addYAxisTitle,
  addLogo,
} from "./chromeShapes";
import { logoBase64 } from "./chromeLogo";

export interface ChromeBuildResult {
  /** Names of the chrome shapes created this run, in z-order (canvas first). */
  createdShapeNames: string[];
  /** Non-fatal issues (e.g. the logo could not be added). */
  warnings: string[];
}

export function buildChrome(
  sheet: Excel.Worksheet,
  chartName: string,
  baseLeft: number,
  baseTop: number,
  defaults: ChartDefaults,
  groupExists: boolean
): ChromeBuildResult {
  removeExistingChrome(sheet, chartName, groupExists);

  const pos = chromePositions(baseLeft, baseTop, defaults.showYAxisTitle);
  const createdShapeNames: string[] = [];
  const warnings: string[] = [];

  // Canvas first so it sits at the back of the z-order.
  createdShapeNames.push(addCanvas(sheet, chartName, pos.canvas).name);
  createdShapeNames.push(addFigureBox(sheet, chartName, pos.figure).name);
  createdShapeNames.push(addTitleBox(sheet, chartName, pos.title).name);
  createdShapeNames.push(addSubtitleBox(sheet, chartName, pos.subtitle).name);
  createdShapeNames.push(addSourceBox(sheet, chartName, pos.source).name);

  const logo = addLogo(sheet, chartName, logoBase64, pos.logo);
  if (logo) {
    createdShapeNames.push(logo.name);
  } else {
    warnings.push("The logo could not be added; the rest of the chrome was built.");
  }

  if (defaults.showYAxisTitle && pos.yAxis) {
    createdShapeNames.push(addYAxisTitle(sheet, chartName, pos.yAxis).name);
  }

  return { createdShapeNames, warnings };
}
