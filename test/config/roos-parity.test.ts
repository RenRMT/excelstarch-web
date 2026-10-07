/**
 * Insurance against a silent transcription error in the palette: read the ROOS colour picker
 * (`reference/roos-kleurenkiezer.html`, the auditable source of every colour) at runtime, extract
 * its token tables, and assert the ported ramps, data colours and brand colours match it exactly.
 */
import { readFileSync } from "fs";
import { resolve } from "path";

import * as brand from "../../src/config/brand";
import { dataColors, colorBrand1 } from "../../src/config/brand";
import { ramps, divergingTags, rampOrder } from "../../src/config/ramps";
import { titleFontColor } from "../../src/config/fonts";

const html = readFileSync(resolve(__dirname, "../../reference/roos-kleurenkiezer.html"), "utf8");

/** The body of `const NAME = <open> … <close>;` in the reference script. */
function block(name: string, open: string, close: string): string {
  const start = html.indexOf(`const ${name} = ${open}`);
  if (start < 0) throw new Error(`${name} not found in the reference file`);
  return html.slice(start, html.indexOf(`${close};`, start));
}

/** Every match of a global regex (the tsconfig lib predates `String.matchAll`). */
function allMatches(text: string, re: RegExp): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) out.push(m);
  return out;
}

/** RAMPS: `Name: ['#…', …]` rows. */
const roosRamps: Record<string, string[]> = {};
allMatches(block("RAMPS", "{", "}"), /(\w+):\s*\[([^\]]+)\]/g).forEach((m) => {
  roosRamps[m[1]] = allMatches(m[2], /#[0-9A-F]{6}/gi).map((h) => h[0].toUpperCase());
});

/** `{ name:'…', value:'#…' }` entries of a token list. */
function tokenList(name: string): { name: string; value: string }[] {
  return allMatches(block(name, "[", "]"), /name:'([^']+)',\s*value:'(#[0-9A-Fa-f]{6})'/g).map(
    (m) => ({ name: m[1], value: m[2].toUpperCase() })
  );
}

const reekspalet = tokenList("REEKSPALET");
const grijs = tokenList("GRIJS");

describe("palette parity with reference/roos-kleurenkiezer.html", () => {
  it("found all nine ramps, the eight-colour series palette and the grey scale", () => {
    expect(Object.keys(roosRamps)).toHaveLength(9);
    expect(reekspalet).toHaveLength(8);
    expect(grijs).toHaveLength(12);
  });

  it("ports every ROOS ramp tint-for-tint, in the same order", () => {
    expect([...rampOrder]).toEqual(Object.keys(roosRamps));
    rampOrder.forEach((name) => {
      expect([...ramps[name]]).toEqual(roosRamps[name]);
    });
  });

  it("data colours are the ROOS series palette in order", () => {
    expect([...dataColors]).toEqual(reekspalet.map((t) => t.value));
  });

  it("chart titles use Lintblauw", () => {
    expect(colorBrand1).toBe(roosRamps.Lintblauw[5]);
    expect(titleFontColor).toBe(colorBrand1);
  });

  it("every colour constant in brand.ts is a ROOS token", () => {
    const tokens = Object.keys(roosRamps)
      .reduce<string[]>((all, name) => all.concat(roosRamps[name]), [])
      .concat(grijs.map((t) => t.value));
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
