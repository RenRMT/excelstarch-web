/**
 * The chart-spec table is the single source of truth for what each chart kind maps to. These tests
 * guard its internal consistency rather than restating its values.
 */
import {
  chartSpecs,
  chartKinds,
  chartFamilies,
  kindForExcelType,
} from "../../src/config/chartSpecs";
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

  it("groups every kind into exactly one family, keeping table order", () => {
    expect(chartFamilies.flatMap((f) => f.kinds)).toEqual(chartKinds);
    expect(chartFamilies.map((f) => f.family)).toEqual([
      "Column",
      "Bar",
      "Line & Area",
      "Pie",
      "Scatter",
      "Complex",
    ]);
  });

  it("looks a kind up by its Excel chart type", () => {
    for (const kind of chartKinds) {
      expect(kindForExcelType(chartSpecs[kind].excelType)).toBe(kind);
    }
    expect(kindForExcelType("Area")).toBeUndefined(); // plain overlapping area is not built
  });

  it.each(chartKinds)("%s: stacked kinds overlap fully, clustered ones do not", (kind) => {
    const { excelType, finishing } = chartSpecs[kind];
    if (/Stacked/.test(excelType) && finishing.overlap !== undefined) {
      expect(finishing.overlap).toBe(100);
    }
    if (/Clustered/.test(excelType)) {
      expect(finishing.overlap).toBeLessThan(100);
    }
  });

  it.each(chartKinds)("%s: round-chart finishing only on round charts", (kind) => {
    const { colourTarget, excelType, finishing } = chartSpecs[kind];
    if (finishing.squarePlot) {
      expect(colourTarget).toBe("point");
    }
    if (finishing.doughnutHoleSize !== undefined) {
      expect(excelType).toMatch(/^Doughnut/);
    }
  });
});

