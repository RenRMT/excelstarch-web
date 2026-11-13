/**
 * Chart-type classification — mirrors the type lists in the VBA `GetStyleColorMode`,
 * `IsLineTarget` and `IsPieChartType`, with markers split out.
 */
import {
  colourTargetForChartType,
  isPieChartType,
  styleClassForChartType,
  isScatterChartType,
} from "../../src/logic/chartClass";

describe("colourTargetForChartType", () => {
  it("colours filled charts on the series fill", () => {
    for (const t of ["ColumnClustered", "BarStacked100", "Area", "AreaStacked", "Bubble", "Boxwhisker"]) {
      expect(colourTargetForChartType(t)).toBe("fill");
    }
  });

  it("colours marker-less line types on the line", () => {
    for (const t of ["Line", "LineStacked", "LineStacked100", "XYScatterLinesNoMarkers", "XYScatterSmoothNoMarkers"]) {
      expect(colourTargetForChartType(t)).toBe("line");
    }
  });

  it("colours line-with-marker types on both", () => {
    for (const t of ["LineMarkers", "LineMarkersStacked", "XYScatterLines", "XYScatterSmooth"]) {
      expect(colourTargetForChartType(t)).toBe("lineMarker");
    }
  });

  it("colours a markers-only scatter on its markers", () => {
    expect(colourTargetForChartType("XYScatter")).toBe("marker");
  });

  it("colours single-series point charts per point", () => {
    for (const t of ["Pie", "Doughnut", "DoughnutExploded", "Treemap", "Sunburst"]) {
      expect(colourTargetForChartType(t)).toBe("point");
    }
  });

  it("falls back to fill for unknown types", () => {
    expect(colourTargetForChartType("SomethingNew")).toBe("fill");
    expect(colourTargetForChartType("")).toBe("fill");
  });
});

describe("isPieChartType", () => {
  it("is true for pie and doughnut only", () => {
    expect(isPieChartType("Pie")).toBe(true);
    expect(isPieChartType("PieExploded")).toBe(true);
    expect(isPieChartType("Doughnut")).toBe(true);
    expect(isPieChartType("DoughnutExploded")).toBe(true);
    expect(isPieChartType("Treemap")).toBe(false);
    expect(isPieChartType("ColumnClustered")).toBe(false);
  });
});

describe("styleClassForChartType", () => {
  it("buckets the VBA classes", () => {
    expect(styleClassForChartType("Pie")).toBe("pie");
    expect(styleClassForChartType("DoughnutExploded")).toBe("pie");
    expect(styleClassForChartType("ColumnStacked")).toBe("axisFill");
    expect(styleClassForChartType("Area")).toBe("axisFill");
    expect(styleClassForChartType("Bubble")).toBe("axisFill");
    expect(styleClassForChartType("LineMarkersStacked")).toBe("lineScatter");
    expect(styleClassForChartType("XYScatterSmooth")).toBe("lineScatter");
  });

  it("sends 3-D, stock, radar, chartex and unknown types to other", () => {
    for (const t of ["3DColumnClustered", "3DLine", "StockHLC", "Radar", "Waterfall", "Sunburst", "Funnel", "Nope"]) {
      expect(styleClassForChartType(t)).toBe("other");
    }
  });
});

describe("isScatterChartType", () => {
  it("is true for the XY scatter family only", () => {
    expect(isScatterChartType("XYScatter")).toBe(true);
    expect(isScatterChartType("XYScatterLinesNoMarkers")).toBe(true);
    expect(isScatterChartType("Line")).toBe(false);
  });
});
