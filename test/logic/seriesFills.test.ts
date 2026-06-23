/**
 * Unit tests for the series-fill ordering (palette / single ramp / diverging / invert + guards).
 * Mirrors the VBA Build*Ramp loops; assertions compose from the config (not magic hex) where the
 * ordering is the point, with a few concrete spot-checks for the documented cases.
 */
import {
  paletteFills,
  singleRampFills,
  divergingFills,
  invertFills,
  MAX_SINGLE_RAMP_SERIES,
  MAX_DIVERGING_SERIES,
} from "../../src/logic/seriesFills";
import { ramps } from "../../src/config/ramps";
import { colorBrand4 } from "../../src/config/brand";
import { getPaletteColor } from "../../src/logic/colorSeries";

describe("paletteFills", () => {
  it("returns count fills in Contrasting order by default", () => {
    expect(paletteFills(3)).toEqual([getPaletteColor(1), getPaletteColor(2), getPaletteColor(3)]);
  });

  it("Rainbow order swaps slot 2 (Lavender) and slot 6 (Coral)", () => {
    const rainbow = paletteFills(8, true);
    // contrasting: data1..8; rainbow reorders [1,7,3,4,5,2,6,8]
    expect(rainbow[1]).toBe(getPaletteColor(2, true)); // slot 2 → colorData7 (Lavender)
    expect(rainbow[5]).toBe(getPaletteColor(6, true)); // slot 6 → colorData2 (Coral)
    expect(rainbow[1]).not.toBe(paletteFills(8, false)[1]);
  });

  it("count 0 yields no fills", () => {
    expect(paletteFills(0)).toEqual([]);
  });
});

describe("singleRampFills", () => {
  it("n=3 of ramp A → steps [6,4,2] darkest-first", () => {
    // orderedRampSteps(3) = [6,4,2]; ramp index is step-1.
    expect(singleRampFills("A", 3)).toEqual([ramps.A[5], ramps.A[3], ramps.A[1]]);
  });

  it("n=1 → the single priority step (6)", () => {
    expect(singleRampFills("A", 1)).toEqual([ramps.A[5]]);
  });

  it("n=10 → all ten steps darkest→lightest", () => {
    expect(singleRampFills("A", 10)).toEqual([
      ramps.A[9],
      ramps.A[8],
      ramps.A[7],
      ramps.A[6],
      ramps.A[5],
      ramps.A[4],
      ramps.A[3],
      ramps.A[2],
      ramps.A[1],
      ramps.A[0],
    ]);
  });

  it("throws when count exceeds the single-ramp limit", () => {
    expect(() => singleRampFills("A", MAX_SINGLE_RAMP_SERIES + 1)).toThrow(RangeError);
  });
});

describe("divergingFills", () => {
  it("n=5 (A|B): left dark→light, grey centre, right light→dark", () => {
    // sideCount=2, priorityStepsSorted(2)=[2,6]. Left desc: steps [6,2]; right asc: steps [2,6].
    expect(divergingFills("A|B", 5)).toEqual([
      ramps.A[5], // left step 6 (dark)
      ramps.A[1], // left step 2 (light)
      colorBrand4, // grey centre (odd count)
      ramps.B[1], // right step 2 (light)
      ramps.B[5], // right step 6 (dark)
    ]);
  });

  it("n=8 (A|B): four per side, no grey centre", () => {
    const fills = divergingFills("A|B", 8);
    expect(fills).toHaveLength(8);
    expect(fills).not.toContain(colorBrand4);
  });

  it("throws on an invalid tag", () => {
    expect(() => divergingFills("A", 4)).toThrow(RangeError);
  });

  it("throws on an unknown ramp", () => {
    expect(() => divergingFills("A|Z", 4)).toThrow(RangeError);
  });

  it("throws when count exceeds the diverging limit", () => {
    expect(() => divergingFills("A|B", MAX_DIVERGING_SERIES + 1)).toThrow(RangeError);
  });
});

describe("invertFills", () => {
  it("reverses the current assignment", () => {
    expect(invertFills(["#111111", "#222222", "#333333"])).toEqual([
      "#333333",
      "#222222",
      "#111111",
    ]);
  });

  it("does not mutate the input", () => {
    const current = ["#111111", "#222222"];
    invertFills(current);
    expect(current).toEqual(["#111111", "#222222"]);
  });
});
