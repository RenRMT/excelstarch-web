/**
 * Style profiles — what the build flow does to a chart, whether it was picked by kind or found
 * selected on the sheet (Apply Chart Style).
 */
import { profileForKind, profileForChartType } from "../../src/logic/styleProfile";
import { chartSpecs, chartKinds } from "../../src/config/chartSpecs";
import { Axis } from "../../src/config/enums";

describe("profileForKind", () => {
  it.each(chartKinds)("%s: takes everything from the spec and styles fully", (kind) => {
    const p = profileForKind(kind);
    expect(p.defaults).toEqual(chartSpecs[kind].defaults());
    expect(p.colourTarget).toBe(chartSpecs[kind].colourTarget);
    expect(p.finishing).toBe(chartSpecs[kind].finishing);
    expect(p.keepAxesAndLegend).toBe(false);
    expect(p.bestEffortColour).toBe(false);
  });
});

describe("profileForChartType", () => {
  it("uses the builder's profile for a type the builder makes", () => {
    expect(profileForChartType("BarStacked")).toEqual(profileForKind("barStacked"));
    expect(profileForChartType("Doughnut")).toEqual(profileForKind("doughnut"));
  });

  it("styles other pies as round charts", () => {
    const p = profileForChartType("PieExploded");
    expect(p.colourTarget).toBe("point");
    expect(p.finishing.squarePlot).toBe(true);
    expect(p.defaults.axisDisplay).toBe(Axis.None);
  });

  it("styles other line and scatter types with their axes and line/marker colour", () => {
    const line = profileForChartType("LineMarkersStacked");
    expect(line.colourTarget).toBe("lineMarker");
    expect(line.defaults.gridlines).toBe(Axis.Y);

    const scatter = profileForChartType("XYScatterSmooth");
    expect(scatter.colourTarget).toBe("lineMarker");
    expect(scatter.defaults.gridlines).toBe(Axis.Both);
  });

  it("styles a plain overlapping area chart as an axis-fill chart", () => {
    const p = profileForChartType("Area");
    expect(p.colourTarget).toBe("fill");
    expect(p.defaults.axisDisplay).toBe(Axis.Both);
    expect(p.keepAxesAndLegend).toBe(false);
  });

  it("leaves axes and legend alone on unknown types and colours them best-effort", () => {
    for (const t of ["Waterfall", "Sunburst", "3DColumnClustered", "StockHLC", "SomethingNew"]) {
      const p = profileForChartType(t);
      expect(p.keepAxesAndLegend).toBe(true);
      expect(p.bestEffortColour).toBe(true);
      expect(p.defaults.axisDisplay).toBe(Axis.None);
      expect(p.defaults.showYAxisTitle).toBe(false);
      expect(p.finishing).toEqual({});
    }
    expect(profileForChartType("Sunburst").colourTarget).toBe("point");
  });
});
