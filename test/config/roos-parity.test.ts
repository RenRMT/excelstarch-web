/**
 * Insurance against a silent transcription error in the palette: read the ROOS reference values
 * (`reference/roos-palette.json`, the auditable source of every colour and the typeface) at runtime
 * and assert the ported ramps, data colours, brand colours and font match it exactly.
 *
 * This reads the file via `fs` rather than importing JSON so the webpack/tsconfig build is
 * unaffected (no `resolveJsonModule` needed).
 */
import { readFileSync } from "fs";
import { resolve } from "path";

import * as brand from "../../src/config/brand";
import { dataColors, colorBrand1 } from "../../src/config/brand";
import { ramps, divergingTags, rampOrder } from "../../src/config/ramps";
import { titleFontColor, fontPrimary } from "../../src/config/fonts";

interface Token {
  name: string;
  value: string;
}

interface RoosPalette {
  font: string;
  ramps: Record<string, string[]>;
  seriesPalette: Token[];
  greys: Token[];
}

const roos: RoosPalette = JSON.parse(
  readFileSync(resolve(__dirname, "../../reference/roos-palette.json"), "utf8")
) as RoosPalette;

describe("palette parity with reference/roos-palette.json", () => {
  it("ports every ROOS ramp tint-for-tint, in the same order", () => {
    expect([...rampOrder]).toEqual(Object.keys(roos.ramps));
    rampOrder.forEach((name) => {
      expect([...ramps[name]]).toEqual(roos.ramps[name]);
    });
  });

  it("data colours are the ROOS series palette in order", () => {
    expect([...dataColors]).toEqual(roos.seriesPalette.map((t) => t.value));
  });

  it("chart titles use Lintblauw", () => {
    expect(colorBrand1).toBe(roos.ramps.Lintblauw[5]);
    expect(titleFontColor).toBe(colorBrand1);
  });

  it("charts use the ROOS typeface", () => {
    expect(fontPrimary).toBe(roos.font);
  });

  it("every colour constant in brand.ts is a ROOS token", () => {
    const tokens = Object.keys(roos.ramps)
      .reduce<string[]>((all, name) => all.concat(roos.ramps[name]), [])
      .concat(roos.greys.map((t) => t.value));
    const exported = brand as unknown as Record<string, unknown>;
    Object.keys(exported)
      .filter((name) => typeof exported[name] === "string")
      .forEach((name) => {
        const value = (exported[name] as string).toUpperCase();
        // The name rides along so a failure says which constant is off-palette.
        expect({ name, isToken: tokens.indexOf(value) >= 0 }).toEqual({ name, isToken: true });
      });
  });

  it("offers every ordered pair of two different ramps as a diverging tag", () => {
    expect(divergingTags).toHaveLength(9 * 8);
    expect(divergingTags).toContain("Rood|Groen");
    expect(divergingTags).not.toContain("Rood|Rood");
  });
});
