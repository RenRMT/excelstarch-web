/**
 * Mirrors `TestPlotAreaGeometry` from `modTestHarness.bas`. Assertions are composed from the
 * named geometry constants (not magic numbers), so the test validates which bands are included
 * per flag combination and stays valid if a proportion is retuned.
 */
import {
  plotAreaTopFor,
  plotAreaHeightFor,
  calcTitlesHeight,
  legendHeight,
  yAxisLabelHeight,
  yAxisLabelPad,
  xAxisLabelHeight,
  plotAreaBottomMargin,
  logoHeight,
  chartHeight,
} from "../../src/config/geometry";

describe("plotAreaTopFor", () => {
  it("top band = titles, plus legend strip and/or y-axis-label strip when shown", () => {
    expect(plotAreaTopFor(false, false)).toBe(calcTitlesHeight);
    expect(plotAreaTopFor(false, true)).toBe(calcTitlesHeight + legendHeight);
    expect(plotAreaTopFor(true, false)).toBe(calcTitlesHeight + yAxisLabelHeight + yAxisLabelPad);
    expect(plotAreaTopFor(true, true)).toBe(
      calcTitlesHeight + legendHeight + yAxisLabelHeight + yAxisLabelPad
    );
  });
});

describe("plotAreaHeightFor", () => {
  it("height = canvas less top band, x-title strip (when shown), and reserved bottom + logo", () => {
    expect(plotAreaHeightFor(true, true, true)).toBeCloseTo(
      chartHeight - plotAreaTopFor(true, true) - xAxisLabelHeight - plotAreaBottomMargin - logoHeight,
      10
    );
    expect(plotAreaHeightFor(false, false, false)).toBeCloseTo(
      chartHeight - plotAreaTopFor(false, false) - plotAreaBottomMargin - logoHeight,
      10
    );
  });

  it("keeps the minimal-chrome plot area positive", () => {
    expect(plotAreaHeightFor(false, false, false)).toBeGreaterThan(0);
  });
});
