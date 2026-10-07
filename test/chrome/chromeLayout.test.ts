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
  titleBoxTop,
  subtitleBoxTop,
  sourceBoxHeight,
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
    expect(p.title.top).toBeCloseTo(base.top, 10); // title sits at the very top
    expect(p.title.top).toBeCloseTo(base.top + titleBoxTop, 10);
    expect(p.subtitle.top).toBeCloseTo(base.top + subtitleBoxTop, 10);

    // Source anchored at the canvas bottom-left.
    expect(p.source.left).toBeCloseTo(base.left, 10);
    expect(p.source.top).toBeCloseTo(base.top + chartHeight - sourceBoxHeight, 10);
  });

  it("builds no figure box or logo", () => {
    const p = chromePositions(0, 0, true);
    expect(Object.keys(p).sort()).toEqual(["canvas", "source", "subtitle", "title", "yAxis"]);
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

  it("ends the chart band (plus x-title strip) exactly at the footer's top edge", () => {
    const p = chromePositions(100, 50, false);
    const noX = chartBandPosition(100, 50, false, false, false);
    expect(noX.top + noX.height).toBeCloseTo(p.source.top, 10);
  });

  it("pie combination (no axes, legend on): top band drops the y-axis strip, keeps the legend", () => {
    // pieChartDefaults → showY=false, showX=false, hasLegend=true. The band must compose from the
    // legend-on/no-y-axis geometry, not the bar/column (axes-on) path.
    const band = chartBandPosition(100, 50, false, false, true);
    expect(band.left).toBeCloseTo(100 + plotAreaLeft, 10);
    expect(band.top).toBeCloseTo(50 + plotAreaTopFor(false, true), 10);
    expect(band.width).toBeCloseTo(chartWidth - 2 * plotAreaLeft, 10);
    expect(band.height).toBeCloseTo(plotAreaHeightFor(false, false, true), 10);
  });
});
