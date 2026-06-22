/**
 * Insurance against a silent transcription error in the ported config: read `reference/spec.json`
 * (the reviewed port input) at runtime and assert the ported brand colours and ramps match it
 * exactly. A mistyped hex digit in a ramp tuple is the most likely invisible bug in the port.
 *
 * This reads the spec via `fs` rather than importing JSON so the webpack/tsconfig build is
 * unaffected (no `resolveJsonModule` needed).
 */
import { readFileSync } from "fs";
import { resolve } from "path";

import {
  colorBrand1,
  colorBrand2,
  colorBrand3,
  colorNeutral2,
  colorNeutral4,
  dataColors,
  orgName,
} from "../../src/config/brand";
import { ramps, divergingTags, paletteOrder } from "../../src/config/ramps";
import type { RampName } from "../../src/config/ramps";

interface Spec {
  orgName: string;
  colors: {
    brand: Record<string, string>;
    neutral: Record<string, string>;
    data: Record<string, string>;
  };
  ramps: Record<string, string[]>;
  divergingTags: string[];
  paletteOrder: { contrasting: number[]; rainbow: number[] };
}

const spec: Spec = JSON.parse(
  readFileSync(resolve(__dirname, "../../reference/spec.json"), "utf8")
) as Spec;

describe("config parity with reference/spec.json", () => {
  it("orgName matches", () => {
    expect(orgName).toBe(spec.orgName);
  });

  it("representative brand/neutral colours match", () => {
    expect(colorBrand1).toBe(spec.colors.brand.colorBrand1);
    expect(colorBrand2).toBe(spec.colors.brand.colorBrand2);
    expect(colorBrand3).toBe(spec.colors.brand.colorBrand3);
    expect(colorNeutral2).toBe(spec.colors.neutral.colorNeutral2);
    expect(colorNeutral4).toBe(spec.colors.neutral.colorNeutral4);
  });

  it("all eight data colours match in order", () => {
    const specData = [1, 2, 3, 4, 5, 6, 7, 8].map((i) => spec.colors.data[`colorData${i}`]);
    expect(dataColors).toEqual(specData);
  });

  it("all eight ramps match step-for-step", () => {
    (Object.keys(ramps) as RampName[]).forEach((name) => {
      expect([...ramps[name]]).toEqual(spec.ramps[name]);
    });
  });

  it("diverging tags and palette order match", () => {
    expect([...divergingTags]).toEqual(spec.divergingTags);
    expect([...paletteOrder.contrasting]).toEqual(spec.paletteOrder.contrasting);
    expect([...paletteOrder.rainbow]).toEqual(spec.paletteOrder.rainbow);
  });
});
