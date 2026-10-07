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

  it("Rainbow order swaps slot 2 and slot 6", () => {
    const rainbow = paletteFills(8, true);
    // contrasting: data1..8; rainbow reorders [1,7,3,4,5,2,6,8]
    expect(rainbow[1]).toBe(getPaletteColor(2, true)); // slot 2 → colorData7
    expect(rainbow[5]).toBe(getPaletteColor(6, true)); // slot 6 → colorData2
    expect(rainbow[1]).not.toBe(paletteFills(8, false)[1]);
  });

  it("count 0 yields no fills", () => {
    expect(paletteFills(0)).toEqual([]);
  });
});

describe("singleRampFills", () => {
  it("n=3 of Hemelblauw → steps [6,4,2] darkest-first", () => {
    // orderedRampSteps(3) = [6,4,2]; ramp index is step-1.
    const r = ramps.Hemelblauw;
    expect(singleRampFills("Hemelblauw", 3)).toEqual([r[5], r[3], r[1]]);
  });

  it("n=1 → the base colour (step 6)", () => {
    expect(singleRampFills("Hemelblauw", 1)).toEqual(["#007BC7"]);
  });

  it("n=6 → all six steps darkest→lightest", () => {
    expect(singleRampFills("Hemelblauw", 6)).toEqual([...ramps.Hemelblauw].reverse());
  });

  it("throws when count exceeds the single-ramp limit", () => {
    expect(MAX_SINGLE_RAMP_SERIES).toBe(6);
    expect(() => singleRampFills("Hemelblauw", MAX_SINGLE_RAMP_SERIES + 1)).toThrow(RangeError);
  });
});

describe("divergingFills", () => {
  it("n=5 (Rood|Groen): left dark→light, grey centre, right light→dark", () => {
    // sideCount=2, priorityStepsSorted(2)=[2,6]. Left desc: steps [6,2]; right asc: steps [2,6].
    expect(divergingFills("Rood|Groen", 5)).toEqual([
      ramps.Rood[5], // left step 6 (dark)
      ramps.Rood[1], // left step 2 (light)
      colorBrand4, // grey centre (odd count)
      ramps.Groen[1], // right step 2 (light)
      ramps.Groen[5], // right step 6 (dark)
    ]);
  });

  it("n=13 (Rood|Groen): every tint on both sides around the centre", () => {
    expect(MAX_DIVERGING_SERIES).toBe(13);
    expect(divergingFills("Rood|Groen", 13)).toEqual([
      ...[...ramps.Rood].reverse(),
      colorBrand4,
      ...ramps.Groen,
    ]);
  });

  it("n=8 (Rood|Groen): four per side, no grey centre", () => {
    const fills = divergingFills("Rood|Groen", 8);
    expect(fills).toHaveLength(8);
    expect(fills).not.toContain(colorBrand4);
  });

  it("throws on an invalid tag", () => {
    expect(() => divergingFills("Rood", 4)).toThrow(RangeError);
  });

  it("throws on an unknown ramp", () => {
    expect(() => divergingFills("Rood|Paars", 4)).toThrow(RangeError);
  });

  it("throws when count exceeds the diverging limit", () => {
    expect(() => divergingFills("Rood|Groen", MAX_DIVERGING_SERIES + 1)).toThrow(RangeError);
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
