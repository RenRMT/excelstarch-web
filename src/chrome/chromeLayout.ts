/**
 * Pure layout maths for the worksheet-chrome overlay — ported from the positioning in
 * `modEngineExChrome.bas` (the `AddChartEx*` shape offsets and `PositionChartExChart`).
 *
 * Worksheet shapes use absolute sheet coordinates, so every chrome element is the canvas origin
 * (`baseLeft`/`baseTop`) plus the canvas-relative geometry constant. The chart OBJECT (not an
 * internal plot area — Office.js `Excel.Chart` has none) is inset into the plot band the same way.
 *
 * PURE: imports only `config/geometry`; no `Excel`/`Office`. This is the one chrome module that is
 * unit-testable, so the offset arithmetic is proven without a host.
 */
import {
  chartWidth,
  chartHeight,
  titleBoxTop,
  titleBoxHeight,
  titleBoxWidth,
  subtitleBoxTop,
  subtitleBoxHeight,
  sourceBoxHeight,
  sourceBoxWidth,
  yAxisLabelTopNoLegend,
  yAxisLabelHeight,
  plotAreaLeft,
  plotAreaTopFor,
  plotAreaHeightFor,
} from "../config/geometry";

/** A positioned rectangle in absolute worksheet points. */
export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Positions for every chrome shape on the canvas (yAxis present only when requested). */
export interface ChromePositions {
  canvas: Box;
  title: Box;
  subtitle: Box;
  source: Box;
  yAxis?: Box;
}

/**
 * All chrome-shape positions for a canvas whose top-left is (baseLeft, baseTop).
 * `showYAxisTitle` adds the optional value-axis title box (chartex types with a value axis; false
 * for bar/column).
 */
export function chromePositions(
  baseLeft: number,
  baseTop: number,
  showYAxisTitle: boolean
): ChromePositions {
  const positions: ChromePositions = {
    // White 600×600 backdrop, sent to back so the chart + chrome render on top.
    canvas: { left: baseLeft, top: baseTop, width: chartWidth, height: chartHeight },
    title: {
      left: baseLeft,
      top: baseTop + titleBoxTop,
      width: titleBoxWidth,
      height: titleBoxHeight,
    },
    subtitle: {
      left: baseLeft,
      top: baseTop + subtitleBoxTop,
      width: titleBoxWidth,
      height: subtitleBoxHeight,
    },
    // Source (footer) box anchored at the canvas bottom-left; the chart band ends at its top edge.
    source: {
      left: baseLeft,
      top: baseTop + chartHeight - sourceBoxHeight,
      width: sourceBoxWidth,
      height: sourceBoxHeight,
    },
  };

  if (showYAxisTitle) {
    positions.yAxis = {
      left: baseLeft,
      top: baseTop + yAxisLabelTopNoLegend,
      width: titleBoxWidth,
      height: yAxisLabelHeight,
    };
  }

  return positions;
}

/**
 * The chart object's position as a band inset into the canvas, leaving the title block above and
 * the source (footer) box below. Ported from `PositionChartExChart`: we set the CHART size/position,
 * not an internal plot area.
 */
export function chartBandPosition(
  baseLeft: number,
  baseTop: number,
  showY: boolean,
  showX: boolean,
  hasLegend: boolean
): Box {
  return {
    left: baseLeft + plotAreaLeft,
    top: baseTop + plotAreaTopFor(showY, hasLegend),
    width: chartWidth - 2 * plotAreaLeft,
    height: plotAreaHeightFor(showY, showX, hasLegend),
  };
}
