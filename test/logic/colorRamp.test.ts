/**
 * Mirrors `TestOrderedRampSteps`, `TestDivergingLayout`, and `TestParseDivergingTag`
 * from `modTestHarness.bas`.
 */
import {
  orderedRampSteps,
  divergingSideCount,
  divergingHasMiddle,
  parseDivergingTag,
} from "../../src/logic/colorRamp";

describe("orderedRampSteps", () => {
  // Priority is [6,2,4,3,5,7,8,1,9,10]; take the first n, sort ascending, reverse (darkest first).
  it("n=1: the single darkest priority step", () => {
    expect(orderedRampSteps(1)).toEqual([6]);
  });

  it("n=2: [6,2] → sort [2,6] → reverse [6,2]", () => {
    expect(orderedRampSteps(2)).toEqual([6, 2]);
  });

  it("n=3: [6,2,4] → sort [2,4,6] → reverse [6,4,2]", () => {
    expect(orderedRampSteps(3)).toEqual([6, 4, 2]);
  });

  it("n=5: [6,2,4,3,5] → sort [2,3,4,5,6] → reverse [6,5,4,3,2]", () => {
    expect(orderedRampSteps(5)).toEqual([6, 5, 4, 3, 2]);
  });

  it("n=10: full set, darkest (10) to lightest (1)", () => {
    expect(orderedRampSteps(10)).toEqual([10, 9, 8, 7, 6, 5, 4, 3, 2, 1]);
  });
});

describe("diverging layout", () => {
  it("even count: equal sides, no grey middle", () => {
    expect(divergingSideCount(8)).toBe(4);
    expect(divergingHasMiddle(8)).toBe(false);
  });

  it("odd count: floor(n/2) per side plus a grey middle", () => {
    expect(divergingSideCount(5)).toBe(2);
    expect(divergingHasMiddle(5)).toBe(true);
  });

  it("single series: no side, but middle pairing (n\\2 = 0)", () => {
    expect(divergingSideCount(1)).toBe(0);
    expect(divergingHasMiddle(1)).toBe(true);
  });
});

describe("parseDivergingTag", () => {
  it("splits a valid 'A|B' into the two ramp names", () => {
    expect(parseDivergingTag("A|B")).toEqual({ left: "A", right: "B" });
  });

  it("rejects a missing separator or empty side", () => {
    expect(parseDivergingTag("A")).toBeNull();
    expect(parseDivergingTag("A|")).toBeNull();
    expect(parseDivergingTag("|B")).toBeNull();
  });
});
