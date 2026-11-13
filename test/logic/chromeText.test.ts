/**
 * Unit tests for the chrome-text resolution rule (blank → placeholder, `-` → empty, text → text).
 */
import { resolveChromeText, splitSourceNotes, fieldValueFor } from "../../src/logic/chromeText";
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

describe("splitSourceNotes", () => {
  it("splits the two lines on any host line-break character", () => {
    expect(splitSourceNotes("Source: a\nNotes: b")).toEqual({ source: "Source: a", notes: "Notes: b" });
    expect(splitSourceNotes("Source: a\rNotes: b")).toEqual({ source: "Source: a", notes: "Notes: b" });
    expect(splitSourceNotes("Source: a\vNotes: b")).toEqual({ source: "Source: a", notes: "Notes: b" });
    expect(splitSourceNotes("Source: a\r\nNotes: b")).toEqual({ source: "Source: a", notes: "Notes: b" });
  });

  it("keeps an empty line empty rather than shifting the other line into it", () => {
    expect(splitSourceNotes("\nNotes: b")).toEqual({ source: "", notes: "Notes: b" });
    expect(splitSourceNotes("Source: a\n")).toEqual({ source: "Source: a", notes: "" });
    expect(splitSourceNotes("Source: a")).toEqual({ source: "Source: a", notes: "" });
  });
});

describe("fieldValueFor (inverse of resolveChromeText)", () => {
  const placeholder = "Title in 28pt sentence case";

  it("shows the placeholder as an empty field, a cleared box as the sentinel, text verbatim", () => {
    expect(fieldValueFor(placeholder, placeholder)).toBe("");
    expect(fieldValueFor("", placeholder)).toBe(clearFieldToken);
    expect(fieldValueFor("My title", placeholder)).toBe("My title");
  });

  it("round-trips: applying the shown value writes back the same text", () => {
    for (const text of [placeholder, "", "My title"]) {
      expect(resolveChromeText(fieldValueFor(text, placeholder), placeholder)).toBe(text);
    }
  });
});
