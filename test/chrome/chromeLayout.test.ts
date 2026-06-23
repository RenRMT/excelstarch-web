/**
 * Unit tests for the pure chrome layout maths. The rest of chrome/ is interop (sideload-tested per
 * docs/testing-manual.md); chromeLayout is pure, so its offset arithmetic is proven here.
 *
 * Assertions are composed from the geometry constants (not magic numbers), so they validate the
 * offset *composition* and stay valid if a proportion is retuned.
 */
import { chromePositions, chartBandPosition } from "../../src/chrome/chromeLayout";
import {
  chartWidth,
  chartHeight,
  figureBoxTop,
  titleBoxTop,
  subtitleBoxTop,
  sourceBoxHeight,
  logoHeight,
  logoAspectRatio,
  logoMarginRight,
  logoMarginBottom,
  yAxisLabelTopNoLegend,
  plotAreaLeft,
  plotAreaTopFor,
  plotAreaHeightFor,
} from "../../src/config/geometry";

describe("chromePositions", () => {
  it("offsets every box from the canvas origin", () => {
    const base = { left: 100, top: 50 };
    const p = chromePositions(base.left, base.top, false);

    expect(p.canvas).toEqual({ left: 100, top: 50, width: chartWidth, height: chartHeight });
    expect(p.figure.top).toBeCloseTo(base.top + figureBoxTop, 10);
    expect(p.title.top).toBeCloseTo(base.top + titleBoxTop, 10);
    expect(p.subtitle.top).toBeCloseTo(base.top + subtitleBoxTop, 10);

    // Source anchored at the canvas bottom-left.
    expect(p.source.left).toBeCloseTo(base.left, 10);
    expect(p.source.top).toBeCloseTo(base.top + chartHeight - sourceBoxHeight, 10);
  });

  it("places the logo bottom-right with the correct width from aspect ratio", () => {
    const base = { left: 0, top: 0 };
    const p = chromePositions(base.left, base.top, false);
    const expectedWidth = logoHeight * logoAspectRatio;

    expect(p.logo.width).toBeCloseTo(expectedWidth, 10);
    expect(p.logo.height).toBeCloseTo(logoHeight, 10);
    expect(p.logo.left).toBeCloseTo(base.left + chartWidth - expectedWidth - logoMarginRight, 10);
    expect(p.logo.top).toBeCloseTo(base.top + chartHeight - logoHeight - logoMarginBottom, 10);
  });

  it("omits the y-axis box unless requested, includes it when asked", () => {
    expect(chromePositions(0, 0, false).yAxis).toBeUndefined();

    const withY = chromePositions(10, 20, true);
    expect(withY.yAxis).toBeDefined();
    expect(withY.yAxis?.top).toBeCloseTo(20 + yAxisLabelTopNoLegend, 10);
  });
});

describe("chartBandPosition", () => {
  it("insets the chart object into the plot band per plotAreaTopFor/HeightFor", () => {
    const band = chartBandPosition(100, 50, true, true, false);
    expect(band.left).toBeCloseTo(100 + plotAreaLeft, 10);
    expect(band.top).toBeCloseTo(50 + plotAreaTopFor(true, false), 10);
    expect(band.width).toBeCloseTo(chartWidth - 2 * plotAreaLeft, 10);
    expect(band.height).toBeCloseTo(plotAreaHeightFor(true, true, false), 10);
  });
});
