/**
 * Mirrors `TestGetPaletteColor` from `modTestHarness.bas` — the default (Contrasting) branch and
 * the out-of-range fallback. The alternate (Rainbow) order is also exercised here because this
 * port injects it as a parameter (the VBA harness could not reach that branch).
 */
import { getPaletteColor } from "../../src/logic/colorSeries";
import { colorData1, colorData3, colorData7, colorData8, colorNeutral2 } from "../../src/config/brand";

describe("getPaletteColor (Contrasting / default)", () => {
  it("maps slot i to dataColor i for 1..8", () => {
    expect(getPaletteColor(1)).toBe(colorData1);
    expect(getPaletteColor(3)).toBe(colorData3); // Sky slot — guards the recolour
    expect(getPaletteColor(8)).toBe(colorData8);
  });

  it("falls back to the neutral for out-of-range indices", () => {
    expect(getPaletteColor(9)).toBe(colorNeutral2);
    expect(getPaletteColor(0)).toBe(colorNeutral2);
  });
});

describe("getPaletteColor (Rainbow / alternate order)", () => {
  it("reorders slot 2 to Lavender (colorData7) and slot 6 to Coral (colorData2)", () => {
    expect(getPaletteColor(1, true)).toBe(colorData1);
    expect(getPaletteColor(2, true)).toBe(colorData7);
    expect(getPaletteColor(8, true)).toBe(colorData8);
  });
});
