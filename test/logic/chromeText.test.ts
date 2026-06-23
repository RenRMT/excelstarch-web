/**
 * Unit tests for the chrome-text resolution rule (blank → placeholder, `-` → empty, text → text).
 */
import { resolveChromeText } from "../../src/logic/chromeText";
import { clearFieldToken } from "../../src/config/text";

describe("resolveChromeText", () => {
  const placeholder = "Title in 28pt sentence case";

  it("blank value restores the placeholder", () => {
    expect(resolveChromeText("", placeholder)).toBe(placeholder);
  });

  it("whitespace-only value restores the placeholder", () => {
    expect(resolveChromeText("   ", placeholder)).toBe(placeholder);
  });

  it("the clear sentinel clears to truly empty", () => {
    expect(resolveChromeText(clearFieldToken, placeholder)).toBe("");
  });

  it("the clear sentinel surrounded by whitespace still clears", () => {
    expect(resolveChromeText("  -  ", placeholder)).toBe("");
  });

  it("ordinary text is written verbatim", () => {
    expect(resolveChromeText("Quarterly revenue", placeholder)).toBe("Quarterly revenue");
  });

  it("text that merely contains a dash is not the sentinel", () => {
    expect(resolveChromeText("Q1-Q4 growth", placeholder)).toBe("Q1-Q4 growth");
  });
});
