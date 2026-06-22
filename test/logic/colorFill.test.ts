/**
 * Mirrors `TestParseFillPayload`, `TestIsRemoveFillPayload`, and `TestColorFromName`
 * from `modTestHarness.bas`. The VBA `-1` "unknown" sentinel is `null` in this port.
 */
import { parseFillPayload, isRemoveFillPayload, colorFromName } from "../../src/logic/colorFill";
import { colorData1, colorData8, colorNeutral4 } from "../../src/config/brand";

describe("parseFillPayload", () => {
  it("name only → opaque (0 transparency)", () => {
    expect(parseFillPayload("DATA1")).toEqual({ name: "DATA1", transparency: 0 });
  });

  it("name + transparency", () => {
    expect(parseFillPayload("DATA1|0.5")).toEqual({ name: "DATA1", transparency: 0.5 });
  });

  it("out-of-range transparency clamps to [0, 1]", () => {
    expect(parseFillPayload("DATA2|1.7")).toEqual({ name: "DATA2", transparency: 1 });
  });

  it("non-numeric transparency falls back to 0", () => {
    expect(parseFillPayload("DATA3|abc")).toEqual({ name: "DATA3", transparency: 0 });
  });
});

describe("isRemoveFillPayload", () => {
  it("recognises NONE/NOFILL/OFF (case-insensitive)", () => {
    expect(isRemoveFillPayload("NONE")).toBe(true);
    expect(isRemoveFillPayload("nofill")).toBe(true);
    expect(isRemoveFillPayload("OFF")).toBe(true);
  });

  it("treats a colour name as not-remove", () => {
    expect(isRemoveFillPayload("DATA1")).toBe(false);
  });
});

describe("colorFromName", () => {
  it("resolves known names (case-insensitive)", () => {
    expect(colorFromName("DATA1")).toBe(colorData1);
    expect(colorFromName("data1")).toBe(colorData1);
    expect(colorFromName("DATA8")).toBe(colorData8);
    expect(colorFromName("NEUTRAL4")).toBe(colorNeutral4);
  });

  it("returns null for an unknown name", () => {
    expect(colorFromName("BOGUS")).toBeNull();
  });
});
