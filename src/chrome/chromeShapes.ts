/* global Excel */
/**
 * Worksheet-shape builders for the chrome overlay — ported from the `AddChartEx*` helpers in
 * `modEngineExChrome.bas`. Each builder creates one shape, names it `<chartName><suffix>`, and
 * styles it; positions come from `chromeLayout`. Text boxes are transparent and borderless (the
 * canvas supplies the white background).
 *
 * These QUEUE shape creation and property writes only — the caller syncs. No `context.sync()` here.
 *
 * INTEROP: calls the Office.js shape API. Only excel/ + chrome/ may do this.
 */
import type { Box } from "./chromeLayout";
import { chromeShapeName } from "./chromeNames";
import { colorWhite } from "../config/brand";
import {
  fontPrimary,
  titleFontSize,
  subtitleFontSize,
  axisFontSize,
  sourceFontSize,
  titleFontColor,
  subtitleFontColor,
  axisFontColor,
} from "../config/fonts";
import {
  titlePlaceholder,
  subtitlePlaceholder,
  yAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../config/text";

function positionShape(shape: Excel.Shape, pos: Box): void {
  shape.left = pos.left;
  shape.top = pos.top;
  shape.width = pos.width;
  shape.height = pos.height;
}

/** A borderless white 600×600 rectangle behind the chart and chrome, sent to the back. */
export function addCanvas(sheet: Excel.Worksheet, chartName: string, pos: Box): Excel.Shape {
  const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
  shape.name = chromeShapeName(chartName, "canvas");
  positionShape(shape, pos);
  shape.fill.setSolidColor(colorWhite);
  shape.lineFormat.visible = false;
  // Chart was created before the canvas, so the canvas lands in front; push it behind so the
  // chart series + chrome render on top of the white backdrop.
  shape.setZOrder(Excel.ShapeZOrder.sendToBack);
  return shape;
}

/** Shared text-box scaffold: transparent fill, no border, name + position + text + font. */
function addTextBoxShape(
  sheet: Excel.Worksheet,
  chartName: string,
  suffix: Parameters<typeof chromeShapeName>[1],
  pos: Box,
  text: string,
  font: { size: number; color: string; bold: boolean; italic: boolean }
): Excel.Shape {
  const shape = sheet.shapes.addTextBox(text);
  shape.name = chromeShapeName(chartName, suffix);
  positionShape(shape, pos);
  shape.fill.clear();
  shape.lineFormat.visible = false;

  const f = shape.textFrame.textRange.font;
  f.name = fontPrimary;
  f.size = font.size;
  f.color = font.color;
  f.bold = font.bold;
  f.italic = font.italic; // cross-platform italic — never a named italic family
  return shape;
}

export function addTitleBox(sheet: Excel.Worksheet, chartName: string, pos: Box): Excel.Shape {
  return addTextBoxShape(sheet, chartName, "title", pos, titlePlaceholder, {
    size: titleFontSize,
    color: titleFontColor,
    bold: true,
    italic: false,
  });
}

export function addSubtitleBox(sheet: Excel.Worksheet, chartName: string, pos: Box): Excel.Shape {
  return addTextBoxShape(sheet, chartName, "subtitle", pos, subtitlePlaceholder, {
    size: subtitleFontSize,
    color: subtitleFontColor,
    bold: false,
    italic: false,
  });
}

/** Source + notes share one box, on two lines (matching the VBA `_SourceBox`). */
export function addSourceBox(sheet: Excel.Worksheet, chartName: string, pos: Box): Excel.Shape {
  return addTextBoxShape(
    sheet,
    chartName,
    "source",
    pos,
    `${sourcePlaceholder}\n${notesPlaceholder}`,
    { size: sourceFontSize, color: axisFontColor, bold: false, italic: false }
  );
}

/** Optional value-axis title (chartex types with a value axis); italic. */
export function addYAxisTitle(sheet: Excel.Worksheet, chartName: string, pos: Box): Excel.Shape {
  return addTextBoxShape(sheet, chartName, "yAxis", pos, yAxisPlaceholder, {
    size: axisFontSize,
    color: axisFontColor,
    bold: false,
    italic: true,
  });
}
