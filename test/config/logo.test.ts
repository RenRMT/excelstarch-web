/**
 * The optional logo's layout: bottom-right of the canvas, clear of the source box.
 */
import { logoBox, logoClearsSourceBox, logoEnabled, logoBase64 } from "../../src/config/logo";
import { chartWidth, chartHeight } from "../../src/config/geometry";

describe("logo layout", () => {
  it("sits inside the bottom-right corner of the canvas", () => {
    const box = logoBox();
    expect(box.left + box.width).toBeLessThanOrEqual(chartWidth);
    expect(box.top + box.height).toBeLessThanOrEqual(chartHeight);
    expect(box.left).toBeGreaterThan(chartWidth / 2);
    expect(box.top).toBeGreaterThan(chartHeight / 2);
  });

  it("does not overlap the source box at the configured sizes", () => {
    expect(logoClearsSourceBox()).toBe(true);
  });

  it("is only enabled together with image data", () => {
    if (logoEnabled) {
      expect(logoBase64.length).toBeGreaterThan(0);
    }
  });
});
