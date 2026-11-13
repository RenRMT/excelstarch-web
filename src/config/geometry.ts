/**
 * Derived chart geometry — originally ported from the VBA `modConfigDerived.bas` (constants) and
 * `modEngineBuilder.bas` (the two plot-area helpers). All values are in points (the Office.js shape
 * unit).
 *
 * The two plot-area helpers (`plotAreaTopFor` / `plotAreaHeightFor`) live here, not in the
 * interop layer, because they are pure (no chart access) and are exercised by the Jest suite;
 * the Phase-1 `excel/` chart builder will consume them.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

// --- Canvas --------------------------------------------------------------------------------
export const chartWidth = 600;
export const chartHeight = 600;

// --- Box proportions -----------------------------------------------------------------------
const titleBoxHeightProportion = 0.07;
const subtitleBoxHeightProportion = 0.05;
const yAxisLabelHeightProportion = 0.04;
const legendHeightProportion = 0.04;
const titleBoxWidthProportion = 1;
const sourceBoxWidthProportion = 0.8;
const sourceBoxHeightProportion = 0.08;

// --- Padding -------------------------------------------------------------------------------
export const plotAreaLeftProportion = 0.005;
/** Flat point pad below the y-axis-label strip (not a proportion). */
export const yAxisLabelPad = 10;

// --- Title area (title at the very top, subtitle below it) ---------------------------------
export const titleBoxTop = 0;
export const titleBoxHeight = chartHeight * titleBoxHeightProportion;
export const subtitleBoxTop = titleBoxTop + titleBoxHeight;
export const subtitleBoxHeight = chartHeight * subtitleBoxHeightProportion;
export const titleBoxWidth = chartWidth * titleBoxWidthProportion;
/** Combined title-band height (title + subtitle); used in plot-area maths. */
export const calcTitlesHeight = titleBoxHeight + subtitleBoxHeight;

// --- Legend strip --------------------------------------------------------------------------
export const legendHeight = chartHeight * legendHeightProportion;

// --- Y-axis label strip --------------------------------------------------------------------
export const yAxisLabelHeight = chartHeight * yAxisLabelHeightProportion;
export const yAxisLabelTopNoLegend = calcTitlesHeight;

// --- Source (footer) box -------------------------------------------------------------------
export const sourceBoxWidth = chartWidth * sourceBoxWidthProportion;
export const sourceBoxHeight = chartHeight * sourceBoxHeightProportion;
/** Top edge of the footer; the plot band ends here. */
export const sourceBoxTop = chartHeight - sourceBoxHeight;

// --- X-axis title strip (mirrors the Y strip; reserved just above the footer) --------------
export const xAxisLabelHeight = yAxisLabelHeight;

// --- Plot area (Top/Height are computed by the helpers below) ------------------------------
export const plotAreaLeft = chartWidth * plotAreaLeftProportion;

// --- Line weights --------------------------------------------------------------------------
export const gridlineWeight = 1;
export const axisLineWeight = 1;
/** Line weight for line-chart series (the VBA `FormatSeriesColors` LINE-mode default). */
export const seriesLineWeight = 2;

// --- Series spacing and markers (from the VBA modConfig "Chart Data settings") -------------
/** Space between bar/column groups, as a percentage of the bar width. */
export const seriesGapWidth = 33;
/** Overlap between clustered series; negative leaves a small gap between them. */
export const seriesOverlap = -5;
/** Stacked bars/columns must be flush. */
export const stackedOverlap = 100;
/** Scatter marker size in points. */
export const scatterMarkerSize = 9;
/** Scatter marker shapes, cycled per series (Office.js `ChartMarkerStyle` values). */
export const scatterMarkerStyles: readonly string[] = ["Circle", "Square", "Triangle", "X"];
/** Doughnut hole size in percent (Excel's own default is 75). */
export const donutHoleSize = 50;

// --- Round charts ----------------------------------------------------------------------------
/** Share of the available space a pie/doughnut plot square takes (the rest is breathing room). */
const squarePlotFill = 0.92;

/**
 * The plot area of a pie/doughnut as a centred square inside a chart object of `width` × `height`
 * points, below a `topReserve` strip (the legend). Coordinates are relative to the chart object.
 * Replaces the VBA's fixed pie plot constants, which assumed a 600 × 600 chart.
 */
export function squarePlotArea(
  width: number,
  height: number,
  topReserve: number
): { left: number; top: number; width: number; height: number } {
  const available = Math.max(0, height - topReserve);
  const size = Math.max(0, Math.min(width, available) * squarePlotFill);
  return {
    left: (width - size) / 2,
    top: topReserve + (available - size) / 2,
    width: size,
    height: size,
  };
}

/**
 * Plot-area Top: the title band, plus the legend strip and/or the y-axis-label strip when shown.
 * Ported from `PlotAreaTopFor` in `modEngineBuilder.bas`.
 */
export function plotAreaTopFor(showY: boolean, hasLegend: boolean): number {
  return (
    calcTitlesHeight +
    (hasLegend ? legendHeight : 0) +
    (showY ? yAxisLabelHeight + yAxisLabelPad : 0)
  );
}

/**
 * Plot-area Height: from the top band down to the footer (source box) top edge, less the x-title
 * strip when shown. Adapted from `PlotAreaHeightFor` in `modEngineBuilder.bas`, which reserved a
 * bottom margin plus a logo band instead.
 */
export function plotAreaHeightFor(showY: boolean, showX: boolean, hasLegend: boolean): number {
  return sourceBoxTop - plotAreaTopFor(showY, hasLegend) - (showX ? xAxisLabelHeight : 0);
}

// --- Annotation boxes (from the VBA modConfig "Chart Actions settings") ---------------------
export const annotationBoxWidth = 120;
export const annotationBoxHeight = 30;
/** Box left = data-label left + this offset. */
export const annotationOffsetX = 8;
/** Box top = data-label top + this offset (negative: the box sits above the point). */
export const annotationOffsetY = -8;
