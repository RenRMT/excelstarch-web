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
import {
  fontPrimary,
  titleFontSize,
  subtitleFontSize,
  figureFontSize,
  axisFontSize,
  sourceFontSize,
  titleFontColor,
  subtitleFontColor,
  figureFontColor,
} from "../../src/config/fonts";
import {
  figurePlaceholder,
  titlePlaceholder,
  subtitlePlaceholder,
  yAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../../src/config/text";

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
  fonts: {
    primary: string;
    sizes: Record<string, number>;
    colors: Record<string, string>;
  };
  placeholders: Record<string, string>;
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

  it("font family and sizes match", () => {
    expect(fontPrimary).toBe(spec.fonts.primary);
    expect(titleFontSize).toBe(spec.fonts.sizes.title);
    expect(subtitleFontSize).toBe(spec.fonts.sizes.subtitle);
    expect(figureFontSize).toBe(spec.fonts.sizes.figure);
    expect(axisFontSize).toBe(spec.fonts.sizes.axis);
    expect(sourceFontSize).toBe(spec.fonts.sizes.source);
  });

  it("font colours resolve to the brand colour their spec name points to", () => {
    // spec.fonts.colors holds NAMES (e.g. "colorBrand1"); the ported colour must equal the hex
    // that name resolves to in spec.colors.brand.
    expect(titleFontColor).toBe(spec.colors.brand[spec.fonts.colors.title]);
    expect(subtitleFontColor).toBe(spec.colors.brand[spec.fonts.colors.subtitle]);
    expect(figureFontColor).toBe(spec.colors.brand[spec.fonts.colors.figure]);
  });

  it("placeholder text matches", () => {
    expect(figurePlaceholder).toBe(spec.placeholders.figure);
    expect(titlePlaceholder).toBe(spec.placeholders.title);
    expect(subtitlePlaceholder).toBe(spec.placeholders.subtitle);
    expect(yAxisPlaceholder).toBe(spec.placeholders.yAxis);
    expect(sourcePlaceholder).toBe(spec.placeholders.source);
    expect(notesPlaceholder).toBe(spec.placeholders.notes);
  });
});
