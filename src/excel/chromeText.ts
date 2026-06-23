/* global Excel */
/**
 * Chart-text panel writer — replaces the VBA in-chart selection-driven text editing. Resolves a
 * chrome shape by its name convention (`<chartName>_<Suffix>`) and writes the user's text to it.
 * No chart needs to be selected (more robust than the old Selection model).
 *
 * Source and Notes share one `_SourceBox` on two lines: writing one field reloads the box, replaces
 * the relevant line, and writes both back. An empty value restores the placeholder.
 *
 * INTEROP: Office.js shape API. UI never throws — returns a typed RunResult.
 */
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import { chromeShapeName, ChromeSuffixKey } from "../chrome/chromeNames";
import {
  titlePlaceholder,
  subtitlePlaceholder,
  figurePlaceholder,
  yAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../config/text";

/** Editable chrome fields exposed by the task pane. */
export type ChromeField = "title" | "subtitle" | "figure" | "yAxis" | "source" | "notes";

/** Map a single-box field to its shape suffix + placeholder. */
const SINGLE_BOX: Record<
  Exclude<ChromeField, "source" | "notes">,
  { suffix: ChromeSuffixKey; placeholder: string }
> = {
  title: { suffix: "title", placeholder: titlePlaceholder },
  subtitle: { suffix: "subtitle", placeholder: subtitlePlaceholder },
  figure: { suffix: "figure", placeholder: figurePlaceholder },
  yAxis: { suffix: "yAxis", placeholder: yAxisPlaceholder },
};

export async function writeChromeText(
  chartName: string,
  field: ChromeField,
  value: string
): Promise<RunResult<void>> {
  if (!shapesSupported()) {
    return unsupported<void>();
  }

  return runExcel<void>(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();

    if (field === "source" || field === "notes") {
      await writeSourceOrNotes(ctx, sheet, chartName, field, value);
      return;
    }

    const { suffix, placeholder } = SINGLE_BOX[field];
    const shape = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, suffix));
    shape.load("isNullObject");
    await ctx.sync();
    if (shape.isNullObject) {
      throw new Error(
        `No "${field}" box found for chart "${chartName}". Recreate the chart first.`
      );
    }
    shape.textFrame.textRange.text = value.length > 0 ? value : placeholder;
    await ctx.sync();
  });
}

/** Source + Notes live on two lines of the shared `_SourceBox`; update one line, keep the other. */
async function writeSourceOrNotes(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chartName: string,
  field: "source" | "notes",
  value: string
): Promise<void> {
  const shape = sheet.shapes.getItemOrNullObject(chromeShapeName(chartName, "source"));
  shape.load("isNullObject");
  const range = shape.textFrame.textRange;
  range.load("text");
  await ctx.sync();
  if (shape.isNullObject) {
    throw new Error(`No source box found for chart "${chartName}". Recreate the chart first.`);
  }

  // Office.js text frames may use \r, \n, or \v as the in-shape line separator depending on host,
  // so split on any run of them rather than just "\n" (else editing one field could wipe the other).
  const [existingSource = sourcePlaceholder, existingNotes = notesPlaceholder] =
    range.text.split(/[\r\n\v]+/);

  const source = field === "source" ? value || sourcePlaceholder : existingSource;
  const notes = field === "notes" ? value || notesPlaceholder : existingNotes;

  range.text = `${source}\n${notes}`;
  await ctx.sync();
}
