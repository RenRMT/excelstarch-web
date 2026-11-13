/**
 * The toggle cycles from modEngineToggles.bas.
 */
import {
  nextGridlines,
  nextAxisLabels,
  nextAxisTitles,
  nextDataLabels,
  dataLabelStateFrom,
  isCenterOnlyLabelType,
  describeAxisPair,
  AxisPair,
} from "../../src/logic/toggleCycles";

const none: AxisPair = { value: false, category: false };
const valueOnly: AxisPair = { value: true, category: false };
const categoryOnly: AxisPair = { value: false, category: true };
const both: AxisPair = { value: true, category: true };

describe("nextGridlines", () => {
  it("cycles none → value → category → both → none", () => {
    expect(nextGridlines(none)).toEqual(valueOnly);
    expect(nextGridlines(valueOnly)).toEqual(categoryOnly);
    expect(nextGridlines(categoryOnly)).toEqual(both);
    expect(nextGridlines(both)).toEqual(none);
  });
});

describe("nextAxisLabels", () => {
  it("cycles none → category → value → both → none", () => {
    expect(nextAxisLabels(none)).toEqual(categoryOnly);
    expect(nextAxisLabels(categoryOnly)).toEqual(valueOnly);
    expect(nextAxisLabels(valueOnly)).toEqual(both);
    expect(nextAxisLabels(both)).toEqual(none);
  });
});

describe("nextAxisTitles", () => {
  it("cycles both → Y → X → none → both when both axes exist", () => {
    expect(nextAxisTitles(both, both)).toEqual(valueOnly);
    expect(nextAxisTitles(valueOnly, both)).toEqual(categoryOnly);
    expect(nextAxisTitles(categoryOnly, both)).toEqual(none);
    expect(nextAxisTitles(none, both)).toEqual(both);
  });

  it("never gets stuck on a no-op step with only a value axis", () => {
    // Every click must change something: Y ↔ none.
    expect(nextAxisTitles(valueOnly, valueOnly)).toEqual(none);
    expect(nextAxisTitles(none, valueOnly)).toEqual(valueOnly);
  });

  it("never gets stuck with only a category axis", () => {
    expect(nextAxisTitles(none, categoryOnly)).toEqual(categoryOnly);
    expect(nextAxisTitles(categoryOnly, categoryOnly)).toEqual(none);
  });

  it("stays off when the chart has no axes", () => {
    expect(nextAxisTitles(none, none)).toEqual(none);
  });
});

describe("nextDataLabels", () => {
  it("cycles none → outside → inside → none", () => {
    expect(nextDataLabels("none", false)).toBe("outside");
    expect(nextDataLabels("outside", false)).toBe("inside");
    expect(nextDataLabels("inside", false)).toBe("none");
  });

  it("switches labels in any other position off", () => {
    expect(nextDataLabels("other", false)).toBe("none");
    expect(nextDataLabels("other", true)).toBe("none");
  });

  it("cycles none ↔ inside on stacked charts", () => {
    expect(nextDataLabels("none", true)).toBe("inside");
    expect(nextDataLabels("inside", true)).toBe("none");
    expect(nextDataLabels("outside", true)).toBe("none");
  });
});

describe("dataLabelStateFrom", () => {
  it("maps Office.js positions", () => {
    expect(dataLabelStateFrom(false, "Center")).toBe("none");
    expect(dataLabelStateFrom(true, "OutsideEnd")).toBe("outside");
    expect(dataLabelStateFrom(true, "Center")).toBe("inside");
    expect(dataLabelStateFrom(true, "BestFit")).toBe("other");
    expect(dataLabelStateFrom(true, undefined)).toBe("other");
  });
});

describe("isCenterOnlyLabelType", () => {
  it("is true for stacked types only", () => {
    expect(isCenterOnlyLabelType("ColumnStacked100")).toBe(true);
    expect(isCenterOnlyLabelType("AreaStacked")).toBe(true);
    expect(isCenterOnlyLabelType("ColumnClustered")).toBe(false);
    expect(isCenterOnlyLabelType("Pie")).toBe(false);
  });
});

describe("describeAxisPair", () => {
  it("names the state", () => {
    expect(describeAxisPair(none, "y", "x")).toBe("none");
    expect(describeAxisPair(valueOnly, "y", "x")).toBe("y");
    expect(describeAxisPair(categoryOnly, "y", "x")).toBe("x");
    expect(describeAxisPair(both, "y", "x")).toBe("both");
  });
});
