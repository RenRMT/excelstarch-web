/* global Excel */
/**
 * Chart-text panel reader/writer — replaces the VBA in-chart selection-driven text editing. Finds
 * a chart's chrome shapes by their name convention (`<chartName>_<Suffix>`) on the chart's own
 * worksheet and reads or writes their text.
 *
 * Source and Notes share one `_SourceBox` on two lines: writing one field reloads the box, replaces
 * the relevant line, and writes both back. The blank/sentinel/text rule lives in the pure
 * `logic/chromeText.resolveChromeText` (blank → placeholder, `-` → truly empty, text → text).
 *
 * INTEROP: Office.js shape API. UI never throws — returns a typed RunResult.
 */
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { chromeShapeName, ChromeTextKey } from "../chrome/chromeNames";
import { readChromeSnapshot } from "../chrome/chromeSnapshot";
import { resolveChromeText, splitSourceNotes } from "../logic/chromeText";
import {
  titlePlaceholder,
  subtitlePlaceholder,
  yAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../config/text";

/** Editable chrome fields exposed by the task pane. */
export type ChromeField = "title" | "subtitle" | "yAxis" | "source" | "notes";

/** Identifies a chart across worksheets (chart names are only unique per sheet). */
export interface ChartRef {
  worksheetId: string;
  chartName: string;
}

/**
 * The chart whose text the panel edits, with the current text of each field. A `null` field means
 * the chart has no such box (e.g. no y-axis title on a pie).
 */
export interface ChromeTarget extends ChartRef {
  texts: Record<ChromeField, string | null>;
}

/** Map a single-box field to its shape key + placeholder. */
const SINGLE_BOX: Record<
  Exclude<ChromeField, "source" | "notes">,
  { key: ChromeTextKey; placeholder: string }
> = {
  title: { key: "title", placeholder: titlePlaceholder },
  subtitle: { key: "subtitle", placeholder: subtitlePlaceholder },
  yAxis: { key: "yAxis", placeholder: yAxisPlaceholder },
};

/**
 * Work out which chart the text panel should edit and read its current text: the chart selected in
 * Excel if there is one, otherwise `fallback` (the chart the panel was already editing, or the one
 * just created). Resolves to `null` when there is no such chart or it has no chrome.
 */
export async function resolveChromeTarget(
  fallback: ChartRef | null
): Promise<RunResult<ChromeTarget | null>> {
  if (!shapesSupported()) {
    return unsupported<ChromeTarget | null>();
  }

  return runExcel<ChromeTarget | null>(async (ctx) => {
    const active = ctx.workbook.getActiveChartOrNullObject();
    active.load("isNullObject, name");
    await ctx.sync();

    let ref: ChartRef;
    let sheet: Excel.Worksheet;
    if (!active.isNullObject) {
      sheet = active.worksheet;
      sheet.load("id");
      await ctx.sync();
      ref = { worksheetId: sheet.id, chartName: active.name };
    } else if (fallback !== null) {
      sheet = ctx.workbook.worksheets.getItemOrNullObject(fallback.worksheetId);
      sheet.load("isNullObject");
      await ctx.sync();
      if (sheet.isNullObject) {
        return null;
      }
      ref = fallback;
    } else {
      return null;
    }

    const snapshot = await readChromeSnapshot(ctx, sheet, ref.chartName);
    if (snapshot.canvasOrigin === null) {
      return null; // not a chart this add-in has styled.
    }

    const { title, subtitle, source, yAxis } = snapshot.texts;
    const sourceNotes = source === undefined ? null : splitSourceNotes(source);
    return {
      ...ref,
      texts: {
        title: title ?? null,
        subtitle: subtitle ?? null,
        yAxis: yAxis ?? null,
        source: sourceNotes ? sourceNotes.source : null,
        notes: sourceNotes ? sourceNotes.notes : null,
      },
    };
  });
}

export async function writeChromeText(
  target: ChartRef,
  field: ChromeField,
  value: string
): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  return runExcel<void>(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.worksheetId);
    const key: ChromeTextKey = field === "source" || field === "notes" ? "source" : field;
    const shape = sheet.shapes.getItemOrNullObject(chromeShapeName(target.chartName, key));
    shape.load("isNullObject");
    await ctx.sync();
    if (shape.isNullObject) {
      throw new Error(
        `No "${field}" box found for chart "${target.chartName}". Restyle the chart first.`
      );
    }

    const range = shape.textFrame.textRange;
    if (field === "source" || field === "notes") {
      // Source + Notes live on two lines of the shared box; update one line, keep the other.
      range.load("text");
      await ctx.sync();
      const existing = splitSourceNotes(range.text);
      const source =
        field === "source" ? resolveChromeText(value, sourcePlaceholder) : existing.source;
      const notes = field === "notes" ? resolveChromeText(value, notesPlaceholder) : existing.notes;
      range.text = `${source}\n${notes}`;
    } else {
      range.text = resolveChromeText(value, SINGLE_BOX[field].placeholder);
    }
    await ctx.sync();
  });
}
