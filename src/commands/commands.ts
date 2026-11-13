/* global Office, window */
/**
 * Ribbon commands — the one-click buttons of the VBA ribbon tab, run from a hidden page (the
 * manifest's FunctionFile). Each command calls the same interop the task panes use and then tells
 * Office it has finished (`event.completed()`).
 *
 * A ribbon command has no surface of its own, so success is silent (as in the VBA, the chart change
 * is the feedback) and a failure opens a small dialog with the message. Commands that need input
 * the ribbon cannot carry — chart text, a series, ramp choices — stay in the task panes; the ribbon
 * re-applies the last palette and ramps chosen there.
 *
 * Not a shared runtime: the two task panes stay separate pages, so a command cannot open a pane
 * with a message. Commands re-read the workbook settings before using them (see persist/settings).
 */
import { chartKinds, ChartKind } from "../config/chartSpecs";
import type { RampName } from "../config/ramps";
import { rampOrder, divergingTags } from "../config/ramps";
import type { RunResult } from "../excel/session";
import { createBrandedChart, styleSelectedChart } from "../excel/chartFlow";
import {
  toggleGridlines,
  toggleAxisLabels,
  toggleAxisTitles,
  toggleLegend,
  toggleDataLabels,
} from "../excel/toggleFlow";
import { addAnnotation } from "../excel/annotationFlow";
import { recolourSeries } from "../excel/colourFlow";
import { readExportLayers, insertExportPicture } from "../excel/exportFlow";
import { composeExport } from "../taskpane/export/compose";
import { EXPORT_SCALES } from "../logic/exportLayout";
import { loadLastUsedColoursFresh, loadExportPrefsFresh } from "../persist/settings";

const DEFAULT_RAMP: RampName = "Hemelblauw";
const DEFAULT_DIVERGING_TAG = "Rood|Groen";
const DEFAULT_EXPORT_SCALE = 2;

/** Show a problem in a small dialog; resolves once the dialog is open (or could not open). */
function showMessage(title: string, body: string): Promise<void> {
  const url = new URL("message.html", window.location.href);
  url.searchParams.set("title", title);
  url.searchParams.set("body", body);
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height: 25, width: 30, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => dialog.close());
        }
        resolve();
      }
    );
  });
}

/**
 * Wrap an action as a ribbon command: run it, report a failure in a dialog, and always complete
 * the event (Office shows the button as busy until then).
 */
function command(
  failTitle: string,
  action: () => Promise<RunResult<unknown>>
): (event: Office.AddinCommands.Event) => Promise<void> {
  return async (event) => {
    try {
      const result = await action();
      if (!result.ok) {
        await showMessage(
          result.kind === "unsupported" ? "Unsupported Excel version" : failTitle,
          result.message
        );
      }
    } catch (error) {
      await showMessage(failTitle, error instanceof Error ? error.message : String(error));
    } finally {
      event.completed();
    }
  };
}

/** Re-apply the palette order last used in the Color Picker. */
async function reapplyPalette(): Promise<RunResult<void>> {
  const last = await loadLastUsedColoursFresh();
  return recolourSeries({ kind: "palette", useAltOrder: last.paletteAltOrder === true });
}

/** Re-apply the single-hue ramp last used in the Color Picker. */
async function reapplyRamp(): Promise<RunResult<void>> {
  const last = await loadLastUsedColoursFresh();
  const rampName = rampOrder.find((r) => r === last.rampName) ?? DEFAULT_RAMP;
  return recolourSeries({ kind: "single", rampName });
}

/** Re-apply the diverging ramp last used in the Color Picker. */
async function reapplyDiverging(): Promise<RunResult<void>> {
  const last = await loadLastUsedColoursFresh();
  const tag =
    last.divergingTag && divergingTags.includes(last.divergingTag)
      ? last.divergingTag
      : DEFAULT_DIVERGING_TAG;
  return recolourSeries({ kind: "diverging", tag });
}

/** Insert the composed chart picture next to the chart (the export route that works everywhere). */
async function exportPicture(): Promise<RunResult<void>> {
  const prefs = await loadExportPrefsFresh();
  const scale = EXPORT_SCALES.find((s) => s === prefs.scale) ?? DEFAULT_EXPORT_SCALE;
  const layers = await readExportLayers(scale);
  if (!layers.ok) {
    return layers;
  }
  const image = await composeExport(layers.value, scale, "png");
  return insertExportPicture(layers.value.worksheetId, layers.value.frame, image.base64);
}

/** Function names as referenced by `<FunctionName>` in manifest.xml. */
export function createCommandName(kind: ChartKind): string {
  return `create_${kind}`;
}

Office.onReady(() => {
  for (const kind of chartKinds) {
    Office.actions.associate(
      createCommandName(kind),
      command("Couldn't create the chart", () => createBrandedChart(kind))
    );
  }
  const associate = (name: string, failTitle: string, action: () => Promise<RunResult<unknown>>) =>
    Office.actions.associate(name, command(failTitle, action));

  associate("applyStyle", "Couldn't style the chart", styleSelectedChart);
  associate("toggleGridlines", "Couldn't change the gridlines", toggleGridlines);
  associate("toggleAxisLabels", "Couldn't change the axis labels", toggleAxisLabels);
  associate("toggleAxisTitles", "Couldn't change the axis titles", toggleAxisTitles);
  associate("toggleLegend", "Couldn't change the legend", toggleLegend);
  associate("toggleDataLabels", "Couldn't change the data labels", () => toggleDataLabels("all"));
  associate("addAnnotation", "Couldn't add the annotation", () => addAnnotation({ kind: "centre" }));
  associate("reapplyPalette", "Couldn't apply the palette", reapplyPalette);
  associate("reapplyRamp", "Couldn't apply the ramp", reapplyRamp);
  associate("reapplyDiverging", "Couldn't apply the diverging ramp", reapplyDiverging);
  associate("invertColours", "Couldn't invert the colours", () =>
    recolourSeries({ kind: "invert" })
  );
  associate("exportPicture", "Couldn't export the chart", exportPicture);
});
