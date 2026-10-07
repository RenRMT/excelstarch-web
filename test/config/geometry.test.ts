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
  sourceBoxTop,
  sourceBoxHeight,
  titleBoxHeight,
  subtitleBoxHeight,
  chartHeight,
} from "../../src/config/geometry";

describe("title band", () => {
  it("is just the title + subtitle (no figure-number box)", () => {
    expect(calcTitlesHeight).toBe(titleBoxHeight + subtitleBoxHeight);
  });
});

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
  it("height = top band down to the footer top, less the x-title strip when shown", () => {
    expect(sourceBoxTop).toBe(chartHeight - sourceBoxHeight);
    expect(plotAreaHeightFor(true, true, true)).toBeCloseTo(
      sourceBoxTop - plotAreaTopFor(true, true) - xAxisLabelHeight,
      10
    );
    expect(plotAreaHeightFor(false, false, false)).toBeCloseTo(
      sourceBoxTop - plotAreaTopFor(false, false),
      10
    );
  });

  it("keeps the minimal-chrome plot area positive", () => {
    expect(plotAreaHeightFor(false, false, false)).toBeGreaterThan(0);
  });
});
