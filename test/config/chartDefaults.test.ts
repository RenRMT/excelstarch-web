/**
 * Per-chart-type formatting defaults — ported from `modConfigDerived.bas`'s `*ChartDefaults`
 * factories. Assertions compose from the named `Axis` enum (not magic numbers) so they validate
 * *intent* (which axes show gridlines / are displayed, legend, value-axis title) and stay valid if
 * the enum's backing numbers ever change.
 *
 * One deliberate divergence from the VBA: VBA `BarChartDefaults` set gridlines to `axisX`, but the
 * TS `barChartDefaults` uses `Axis.Y` (the value axis). A horizontal bar rotates the axes, so
 * value-axis gridlines render as the expected vertical lines — see the comment in chartDefaults.ts.
 * These tests assert the corrected TS behaviour.
 */
import {
  defaultChartDefaults,
  barChartDefaults,
  columnChartDefaults,
  lineChartDefaults,
  areaChartDefaults,
  scatterChartDefaults,
  pieChartDefaults,
} from "../../src/config/chartDefaults";
import { Axis, defaultGridlines, defaultAxisDisplay, defaultLegend } from "../../src/config/enums";

describe("defaultChartDefaults (global, minimal chrome)", () => {
  it("uses the spec global defaults and no value-axis title", () => {
    expect(defaultChartDefaults()).toEqual({
      gridlines: defaultGridlines,
      axisDisplay: defaultAxisDisplay,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });
});

describe("classic chart-type defaults", () => {
  it("bar: value-axis (Y) gridlines, both axes, default legend", () => {
    // Axis.Y despite VBA axisX — a horizontal bar rotates the value axis to run horizontally.
    expect(barChartDefaults()).toEqual({
      gridlines: Axis.Y,
      axisDisplay: Axis.Both,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });

  it("column: Y gridlines, both axes, default legend", () => {
    expect(columnChartDefaults()).toEqual({
      gridlines: Axis.Y,
      axisDisplay: Axis.Both,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });

  it("line: Y gridlines, both axes, default legend", () => {
    expect(lineChartDefaults()).toEqual({
      gridlines: Axis.Y,
      axisDisplay: Axis.Both,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });

  it("area: Y gridlines, both axes, default legend", () => {
    expect(areaChartDefaults()).toEqual({
      gridlines: Axis.Y,
      axisDisplay: Axis.Both,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });

  it("scatter: both gridlines (reference grid), both axes, default legend", () => {
    expect(scatterChartDefaults()).toEqual({
      gridlines: Axis.Both,
      axisDisplay: Axis.Both,
      legend: defaultLegend,
      showYAxisTitle: false,
    });
  });

  it("pie: no gridlines, no axes, legend ON (for slice labels)", () => {
    expect(pieChartDefaults()).toEqual({
      gridlines: Axis.None,
      axisDisplay: Axis.None,
      legend: true,
      showYAxisTitle: false,
    });
  });
});

describe("classic factories never request the chartex value-axis title box", () => {
  const classics = [
    barChartDefaults,
    columnChartDefaults,
    lineChartDefaults,
    areaChartDefaults,
    scatterChartDefaults,
    pieChartDefaults,
  ];
  it.each(classics.map((f) => [f.name, f]))("%s sets showYAxisTitle false", (_name, factory) => {
    expect((factory as () => { showYAxisTitle: boolean })().showYAxisTitle).toBe(false);
  });
});
