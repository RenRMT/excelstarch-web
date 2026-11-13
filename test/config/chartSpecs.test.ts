/**
 * The chart-spec table is the single source of truth for what each chart kind maps to. These tests
 * guard its internal consistency rather than restating its values.
 */
import { chartSpecs, chartKinds } from "../../src/config/chartSpecs";
import { colourTargetForChartType } from "../../src/logic/chartClass";

describe("chartSpecs", () => {
  it("lists every kind exactly once, in table order", () => {
    expect(chartKinds).toEqual(Object.keys(chartSpecs));
    expect(new Set(chartKinds).size).toBe(chartKinds.length);
  });

  it.each(chartKinds)("%s: gridlines are a subset of the displayed axes", (kind) => {
    const d = chartSpecs[kind].defaults();
    // A chart can only show gridlines on an axis it displays (see ChartDefaults.gridlines).
    expect(d.gridlines & ~d.axisDisplay).toBe(0);
  });

  it.each(chartKinds)("%s: a value-axis title is only requested when the value axis shows", (kind) => {
    const d = chartSpecs[kind].defaults();
    if (d.showYAxisTitle) {
      expect(d.axisDisplay & 2).not.toBe(0);
    }
  });

  it.each(chartKinds)("%s: colour target agrees with the chart-type classifier", (kind) => {
    // The builder colours by spec; the Color Picker colours by classifying the chart's type. They
    // must agree, or a chart would be recoloured somewhere other than where it was first coloured.
    const spec = chartSpecs[kind];
    expect(colourTargetForChartType(spec.excelType)).toBe(spec.colourTarget);
  });

  it("has unique Excel chart types and labels", () => {
    const types = chartKinds.map((k) => chartSpecs[k].excelType);
    const labels = chartKinds.map((k) => chartSpecs[k].label);
    expect(new Set(types).size).toBe(types.length);
    expect(new Set(labels).size).toBe(labels.length);
  });
});
