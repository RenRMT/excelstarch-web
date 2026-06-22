/**
 * Mirrors `TestRelativeLuminance` and `TestContrastColorForFill` from `modTestHarness.bas`.
 */
import { relativeLuminance, contrastColorForFill } from "../../src/logic/colorContrast";
import { colorWhite, colorBrand3 } from "../../src/config/brand";

describe("relativeLuminance", () => {
  it("maps black and white to the exact WCAG endpoints", () => {
    expect(relativeLuminance("#000000")).toBe(0);
    expect(relativeLuminance("#FFFFFF")).toBe(1);
  });

  it("lands a mid-tone strictly between the endpoints", () => {
    const mid = relativeLuminance("#0077BB"); // arbitrary mid-tone blue
    expect(mid).toBeGreaterThan(0);
    expect(mid).toBeLessThan(1);
  });

  it("weights green far more heavily than blue", () => {
    expect(relativeLuminance("#00FF00")).toBeGreaterThan(relativeLuminance("#0000FF"));
  });
});

describe("contrastColorForFill", () => {
  it("returns white text on dark fills and dark brand text on light fills", () => {
    expect(contrastColorForFill("#000000")).toBe(colorWhite);
    expect(contrastColorForFill("#FFFFFF")).toBe(colorBrand3);
  });
});
