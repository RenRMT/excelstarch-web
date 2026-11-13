/* global document, Image, HTMLCanvasElement, CanvasRenderingContext2D, HTMLImageElement */
/**
 * Compose the export image in the task pane: a white frame, the chrome (canvas + text) and the
 * chart on top. Office.js cannot produce this image itself, because the chart and its chrome are
 * separate objects — `excel/exportFlow` supplies the two layers and their geometry, and
 * `logic/exportLayout` says where each goes.
 *
 * The chrome layer is normally Excel's own rendering of the chrome group. If the host could not
 * render it, the text is drawn here from the box contents and the configured fonts — close to, but
 * not pixel-identical with, what Excel shows.
 *
 * DOM code: runs in the task pane's browser, not in the pure or interop layers.
 */
import type { ExportLayers } from "../../excel/exportFlow";
import { exportLayout, wrapText, ExportFormat, Rect } from "../../logic/exportLayout";
import {
  titleBoxTop,
  titleBoxWidth,
  subtitleBoxTop,
  sourceBoxTop,
  sourceBoxWidth,
  yAxisLabelTopNoLegend,
} from "../../config/geometry";
import {
  fontPrimary,
  titleFontSize,
  subtitleFontSize,
  axisFontSize,
  sourceFontSize,
  titleFontColor,
  subtitleFontColor,
  axisFontColor,
  sourceFontColor,
} from "../../config/fonts";

export interface ComposedImage {
  /** Bare base64 image data (no `data:` prefix). */
  base64: string;
  mimeType: "image/png" | "image/jpeg";
  width: number;
  height: number;
}

const JPEG_QUALITY = 0.92;
/** Default inner margins of an Excel text box, in points. */
const TEXT_MARGIN_LEFT = 7.2;
const TEXT_MARGIN_TOP = 3.6;
const LINE_HEIGHT = 1.2;

function loadImage(base64: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("The chart image could not be read."));
    image.src = `data:image/png;base64,${base64}`;
  });
}

function drawInto(context: CanvasRenderingContext2D, image: HTMLImageElement, rect: Rect): void {
  context.drawImage(image, rect.left, rect.top, rect.width, rect.height);
}

interface TextStyle {
  size: number;
  color: string;
  bold?: boolean;
  italic?: boolean;
}

/** Draw one chrome text box: top-left anchored, wrapped to the box width. `k` is px per point. */
function drawTextBox(
  context: CanvasRenderingContext2D,
  text: string | undefined,
  top: number,
  width: number,
  style: TextStyle,
  k: number
): void {
  if (!text) return;
  const sizePx = style.size * k;
  const slant = style.italic ? "italic " : "";
  const weight = style.bold ? "bold " : "";
  context.font = `${slant}${weight}${sizePx}px ${fontPrimary}, sans-serif`;
  context.fillStyle = style.color;
  context.textBaseline = "top";
  const maxWidth = (width - 2 * TEXT_MARGIN_LEFT) * k;
  const lines = wrapText(text, maxWidth, (s) => context.measureText(s).width);
  lines.forEach((line, i) => {
    const y = (top + TEXT_MARGIN_TOP) * k + i * sizePx * LINE_HEIGHT;
    context.fillText(line, TEXT_MARGIN_LEFT * k, y);
  });
}

/** Fallback chrome: draw the text boxes ourselves at their canvas-relative positions. */
function drawChromeText(context: CanvasRenderingContext2D, layers: ExportLayers, k: number): void {
  const { title, subtitle, source, yAxis } = layers.texts;
  const titleStyle = { size: titleFontSize, color: titleFontColor, bold: true };
  const subtitleStyle = { size: subtitleFontSize, color: subtitleFontColor };
  const yAxisStyle = { size: axisFontSize, color: axisFontColor, italic: true };
  const sourceStyle = { size: sourceFontSize, color: sourceFontColor };
  drawTextBox(context, title, titleBoxTop, titleBoxWidth, titleStyle, k);
  drawTextBox(context, subtitle, subtitleBoxTop, titleBoxWidth, subtitleStyle, k);
  drawTextBox(context, yAxis, yAxisLabelTopNoLegend, titleBoxWidth, yAxisStyle, k);
  drawTextBox(context, source, sourceBoxTop, sourceBoxWidth, sourceStyle, k);
}

/** Compose the layers into one image at `scale`, in the requested format. */
export async function composeExport(
  layers: ExportLayers,
  scale: number,
  format: ExportFormat
): Promise<ComposedImage> {
  const layout = exportLayout(layers.frame, layers.chart, layers.chrome, scale);

  const canvas: HTMLCanvasElement = document.createElement("canvas");
  canvas.width = layout.width;
  canvas.height = layout.height;
  const context = canvas.getContext("2d");
  if (context === null) {
    throw new Error("This browser cannot compose the image (no 2D canvas).");
  }

  // White backdrop first: the chart area is transparent and JPEG has no alpha channel.
  context.fillStyle = "#FFFFFF";
  context.fillRect(0, 0, layout.width, layout.height);

  if (layout.chrome !== null) {
    if (layers.chromeImage !== null) {
      drawInto(context, await loadImage(layers.chromeImage), layout.chrome);
    } else {
      drawChromeText(context, layers, layout.pxPerPoint);
    }
  }
  drawInto(context, await loadImage(layers.chartImage), layout.chart);

  const mimeType = format === "jpeg" ? "image/jpeg" : "image/png";
  const dataUrl = canvas.toDataURL(mimeType, JPEG_QUALITY);
  return {
    base64: dataUrl.slice(dataUrl.indexOf(",") + 1),
    mimeType,
    width: layout.width,
    height: layout.height,
  };
}
