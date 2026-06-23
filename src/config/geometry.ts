/**
 * Derived chart geometry — ported from `modConfigDerived.bas` (constants) and
 * `modEngineBuilder.bas` (the two plot-area helpers), computed from the proportions in
 * `reference/spec.json` (`geometry.*`). All values are in points (the Office.js shape unit),
 * so the geometry transfers from VBA with no unit conversion.
 *
 * The two plot-area helpers (`plotAreaTopFor` / `plotAreaHeightFor`) live here, not in the
 * interop layer, because they are pure (no chart access) and are exercised by the Jest suite;
 * the Phase-1 `excel/` chart builder will consume them.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */

// --- Canvas (spec.geometry.canvas) ---------------------------------------------------------
export const chartWidth = 600;
export const chartHeight = 600;

// --- Box proportions (spec.geometry.boxProportions) ----------------------------------------
const figureBoxHeightProportion = 0.04;
const titleBoxHeightProportion = 0.07;
const subtitleBoxHeightProportion = 0.05;
const yAxisLabelHeightProportion = 0.04;
const legendHeightProportion = 0.04;
const titleBoxWidthProportion = 1;
const titleBoxNudgeProportion = 0;
const sourceBoxWidthProportion = 0.8;
const sourceBoxHeightProportion = 0.08;

// --- Padding (spec.geometry.padding) -------------------------------------------------------
const legendLeftPadProportion = 0;
export const legendLeftPad = chartWidth * legendLeftPadProportion;
export const plotAreaLeftProportion = 0.005;
/** Flat point pad below the y-axis-label strip (not a proportion). */
export const yAxisLabelPad = 10;

// --- Logo geometry (spec.geometry.logo) ----------------------------------------------------
const logoHeightScale = 0.1;
const logoMarginRightProp = 0.01;
const logoMarginBottomProp = 0.01;
export const logoAspectRatio = 2.08;
export const logoHeight = chartHeight * logoHeightScale;
export const logoMarginBottom = chartHeight * logoMarginBottomProp;
export const logoTop = chartHeight - logoHeight - logoMarginBottom;
export const logoMarginRight = chartWidth * logoMarginRightProp;

// --- Plot-area bottom margin (spec.geometry.plotAreaBottomMarginProp) ----------------------
const plotAreaBottomMarginProp = 0.03;
export const plotAreaBottomMargin = chartHeight * plotAreaBottomMarginProp;

// --- Title area (modConfigDerived: figure / title / subtitle bands) ------------------------
export const figureBoxTop = 0;
export const figureBoxHeight = chartHeight * figureBoxHeightProportion;
export const titleBoxTop = figureBoxHeight;
export const titleBoxHeight = chartHeight * titleBoxHeightProportion;
export const subtitleBoxTop = titleBoxTop + titleBoxHeight;
export const subtitleBoxHeight = chartHeight * subtitleBoxHeightProportion;
export const titleBoxWidth = chartWidth * titleBoxWidthProportion;
export const titleBoxNudge = chartWidth * titleBoxNudgeProportion;
/** Combined title-band height (figure + title + subtitle); used in plot-area maths. */
export const calcTitlesHeight = figureBoxHeight + titleBoxHeight + subtitleBoxHeight;

// --- Legend strip --------------------------------------------------------------------------
export const legendTop = calcTitlesHeight;
export const legendHeight = chartHeight * legendHeightProportion;

// --- Y-axis label strip --------------------------------------------------------------------
export const yAxisLabelTop = calcTitlesHeight + legendHeight;
export const yAxisLabelHeight = chartHeight * yAxisLabelHeightProportion;
export const yAxisLabelTopNoLegend = calcTitlesHeight;

// --- X-axis title strip (mirrors the Y strip; sits just above the logo/source band) --------
export const xAxisLabelHeight = yAxisLabelHeight;
export const xAxisLabelTop = chartHeight - logoHeight - plotAreaBottomMargin - xAxisLabelHeight;

// --- Plot area (fixed dimensions; Top/Height are computed by the helpers below) ------------
export const plotAreaWidth = chartWidth;
export const plotAreaLeft = chartWidth * plotAreaLeftProportion;

// --- Source box ----------------------------------------------------------------------------
export const sourceBoxWidth = chartWidth * sourceBoxWidthProportion;
export const sourceBoxHeight = chartHeight * sourceBoxHeightProportion;

// --- Pie / donut geometry (spec.geometry.pie) ----------------------------------------------
export const pieLegendGap = 6;
export const pieLegendTop = subtitleBoxTop + subtitleBoxHeight + pieLegendGap;
export const piePlotAreaSizeLegend = 400;
export const piePlotAreaSizeNoLegend = 447;
export const piePlotAreaLeft = 131;
export const piePlotAreaTop = 53;
export const piePlotTopRatio = 0.75;

// --- Series spacing (spec.geometry.series) -------------------------------------------------
export const seriesGapWidth = 33;
export const seriesOverlap = -5;

// --- Lollipop (spec.geometry.lollipop) -----------------------------------------------------
export const lollipopGapWidth = 150;
export const lollipopStickWeight = 2;

// --- Line weights (spec.geometry.weights) --------------------------------------------------
export const gridlineWeight = 1;
export const axisLineWeight = 1;

// --- Annotation box (spec.geometry.annotation) ---------------------------------------------
export const annotationBoxWidth = 120;
export const annotationBoxHeight = 30;
export const annotationOffsetX = 8;
export const annotationOffsetY = -8;

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
 * Plot-area Height: the canvas below the top band, less the x-title strip (when shown) and the
 * always-reserved bottom margin + logo. Ported from `PlotAreaHeightFor` in `modEngineBuilder.bas`.
 */
export function plotAreaHeightFor(showY: boolean, showX: boolean, hasLegend: boolean): number {
  return (
    chartHeight -
    plotAreaTopFor(showY, hasLegend) -
    (showX ? xAxisLabelHeight : 0) -
    plotAreaBottomMargin -
    logoHeight
  );
}
