/* global Excel */
/**
 * Add an annotation box to the selected chart — the VBA `AnnotateButton`. The box is a worksheet
 * text box (a chart cannot contain shapes in Office.js) placed over the chart: at the centre of the
 * plot, or beside a chosen data point. The VBA used the clicked point; add-ins cannot see clicks,
 * so the pane passes the series and point.
 *
 * Annotations stack (each gets the next free number) and stay out of the chrome group so they
 * remain in front of the chart. Export picks them up as their own layer.
 *
 * INTEROP: Office.js. UI never throws — returns a typed RunResult.
 */
import { RunResult, runExcel, shapesSupported, unsupported } from "./session";
import {
  nextAnnotationName,
  annotationBoxAtCentre,
  annotationBoxAtPoint,
} from "../logic/annotation";
import type { Rect } from "../logic/exportLayout";
import { fontPrimary, axisFontSize, axisFontColor } from "../config/fonts";
import { annotationPlaceholder } from "../config/text";

/** Where to put the box: the plot centre, or beside point `point` of series `series` (0-based). */
export type AnnotationAnchor = { kind: "centre" } | { kind: "point"; series: number; point: number };

const NO_CHART_SELECTED = "No chart selected. Select a chart on the sheet, then try again.";

/** Chart-relative top-left of a point's data label, borrowing a label if it has none. */
async function readPointLabelPosition(
  ctx: Excel.RequestContext,
  point: Excel.ChartPoint
): Promise<{ left: number; top: number } | null> {
  point.load("hasDataLabel");
  await ctx.sync();
  const hadLabel = point.hasDataLabel;
  try {
    if (!hadLabel) {
      point.hasDataLabel = true;
    }
    point.dataLabel.load("left, top");
    await ctx.sync();
    return { left: point.dataLabel.left, top: point.dataLabel.top };
  } catch {
    return null; // e.g. an unplotted point
  } finally {
    if (!hadLabel) {
      // Restore: we only borrowed the label to read its position.
      point.hasDataLabel = false;
      await ctx.sync().catch(() => undefined);
    }
  }
}

/** The plot's inner area in worksheet points, or the chart's own area when it cannot be read. */
async function readPlotArea(ctx: Excel.RequestContext, chart: Excel.Chart): Promise<Rect> {
  const chartRect = { left: chart.left, top: chart.top, width: chart.width, height: chart.height };
  try {
    const plot = chart.plotArea;
    plot.load("insideLeft, insideTop, insideWidth, insideHeight");
    await ctx.sync();
    return {
      left: chart.left + plot.insideLeft,
      top: chart.top + plot.insideTop,
      width: plot.insideWidth,
      height: plot.insideHeight,
    };
  } catch {
    return chartRect; // treemap-family charts may not report a plot area
  }
}

export async function addAnnotation(anchor: AnnotationAnchor): Promise<RunResult<string>> {
  if (!shapesSupported()) {
    return unsupported<string>();
  }

  return runExcel<string>(async (ctx) => {
    const chart = ctx.workbook.getActiveChartOrNullObject();
    chart.load("isNullObject");
    await ctx.sync();
    if (chart.isNullObject) {
      throw new Error(NO_CHART_SELECTED);
    }
    chart.load("name, left, top, width, height");
    const sheet = chart.worksheet;
    sheet.shapes.load("items/name");
    await ctx.sync();

    let box: Rect | null = null;
    let placed = "at the plot centre";
    if (anchor.kind === "point") {
      const point = chart.series.getItemAt(anchor.series).points.getItemAt(anchor.point);
      const label = await readPointLabelPosition(ctx, point);
      if (label !== null) {
        box = annotationBoxAtPoint(chart.left + label.left, chart.top + label.top);
        placed = "beside the point";
      } else {
        placed = "at the plot centre (the point has no position)";
      }
    }
    if (box === null) {
      box = annotationBoxAtCentre(await readPlotArea(ctx, chart));
    }

    const name = nextAnnotationName(
      chart.name,
      sheet.shapes.items.map((s) => s.name)
    );
    const shape = sheet.shapes.addTextBox(annotationPlaceholder);
    shape.name = name;
    shape.left = box.left;
    shape.top = box.top;
    shape.width = box.width;
    shape.height = box.height;
    shape.fill.clear();
    shape.lineFormat.visible = false;
    const font = shape.textFrame.textRange.font;
    font.name = fontPrimary;
    font.size = axisFontSize;
    font.color = axisFontColor;
    shape.setZOrder("BringToFront");
    await ctx.sync();

    return `Added ${name.slice(chart.name.length + 1)} ${placed}. Click it on the sheet to edit the text.`;
  });
}
