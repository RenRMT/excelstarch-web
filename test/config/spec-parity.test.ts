/**
 * Insurance against a silent transcription error in the ported config: read `reference/spec.json`
 * (the reviewed port input) at runtime and assert the ported font sizes, placeholders and palette
 * order still match it. Colours, ramps and the typeface no longer come from the VBA spec — they are
 * checked against the ROOS reference in `roos-parity.test.ts`.
 *
 * This reads the spec via `fs` rather than importing JSON so the webpack/tsconfig build is
 * unaffected (no `resolveJsonModule` needed).
 */
import { readFileSync } from "fs";
import { resolve } from "path";

import { paletteOrder } from "../../src/config/ramps";
import {
  titleFontSize,
  subtitleFontSize,
  axisFontSize,
  sourceFontSize,
} from "../../src/config/fonts";
import {
  titlePlaceholder,
  subtitlePlaceholder,
  yAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../../src/config/text";

interface Spec {
  paletteOrder: { contrasting: number[]; rainbow: number[] };
  fonts: {
    sizes: Record<string, number>;
  };
  placeholders: Record<string, string>;
}

const spec: Spec = JSON.parse(
  readFileSync(resolve(__dirname, "../../reference/spec.json"), "utf8")
) as Spec;

describe("config parity with reference/spec.json", () => {
  it("palette order matches", () => {
    expect([...paletteOrder.contrasting]).toEqual(spec.paletteOrder.contrasting);
    expect([...paletteOrder.rainbow]).toEqual(spec.paletteOrder.rainbow);
  });

  it("font sizes match", () => {
    expect(titleFontSize).toBe(spec.fonts.sizes.title);
    expect(subtitleFontSize).toBe(spec.fonts.sizes.subtitle);
    expect(axisFontSize).toBe(spec.fonts.sizes.axis);
    expect(sourceFontSize).toBe(spec.fonts.sizes.source);
  });

  it("placeholder text matches", () => {
    expect(titlePlaceholder).toBe(spec.placeholders.title);
    expect(subtitlePlaceholder).toBe(spec.placeholders.subtitle);
    expect(yAxisPlaceholder).toBe(spec.placeholders.yAxis);
    expect(sourcePlaceholder).toBe(spec.placeholders.source);
    expect(notesPlaceholder).toBe(spec.placeholders.notes);
  });
});
